import JSZip from 'jszip'
import { readXmlFile } from './readXmlFile'
import { getBorder } from './border'
import { getSlideBackgroundFill, getShapeFill, getSolidFill, getPicFill, getPicFilters } from './fill'
import { getChartInfo } from './chart'
import { getVerticalAlign, getTextAutoFit } from './align'
import { getPosition, getSize } from './position'
import { genTextBody } from './text'
import { getCustomShapePath } from './shape'
import { extractFileExtension, base64ArrayBuffer, getTextByPathList, angleToDegrees, getMimeType, isVideoLink, escapeHtml, hasValidText, numberToFixed } from './utils'
import { getShadow } from './shadow'
import { getTableBorders, getTableCellParams, getTableRowParams } from './table'
import { RATIO_EMUs_Points } from './constants'
import { findOMath, latexFormart, parseOMath } from './math'
import { getShapePath } from './shapePath'
import { parseTransition, findTransitionNode } from './animation'

function getPlaceholderType(node) {
  const ph =
    getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph']) ||
    getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'p:ph']) ||
    getTextByPathList(node, ['p:nvGraphicFramePr', 'p:nvPr', 'p:ph'])
  if (!ph) return null
  const t = getTextByPathList(ph, ['attrs', 'type'])
  return t ? String(t) : null
}

function isHeaderFooterPlaceholderType(phType) {
  return phType === 'dt' || phType === 'sldNum' || phType === 'ftr' || phType === 'hdr'
}

function getNodeName(node) {
  return (
    getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr', 'attrs', 'name']) ||
    getTextByPathList(node, ['p:nvPicPr', 'p:cNvPr', 'attrs', 'name']) ||
    getTextByPathList(node, ['p:nvGraphicFramePr', 'p:cNvPr', 'attrs', 'name']) ||
    ''
  )
}

function isLikelyHeaderFooterName(name) {
  if (!name) return false
  const n = String(name)
  return /(^|\b)(Footer Text|Header Text|Slide Number|Date)(\b|$)/i.test(n) || /页脚|页眉|页码|日期/.test(n)
}

function stripHtmlToPlainText(html) {
  return String(html || '').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim()
}

function isPlaceholderPromptText(plainText) {
  const t = String(plainText || '').replace(/\s+/g, ' ').trim()
  if (!t) return false
  const normalized = t.toLowerCase()
  return (
    normalized === 'click to add title' ||
    normalized === 'click to add text' ||
    normalized === 'click to add subtitle' ||
    normalized === 'click to add notes' ||
    /edit\s+master/.test(normalized) ||
    /edit\s+the\s+master/.test(normalized) ||
    /edit\s+master\s+(title|text)/.test(normalized) ||
    t === '此处添加标题' ||
    t === '单击以添加标题' ||
    t === '单击此处添加标题' ||
    t === '此处添加文本' ||
    t === '单击以添加文本' ||
    t === '单击此处添加文本' ||
    t === '单击以添加副标题' ||
    t === '单击此处添加副标题' ||
    t.includes('编辑母版')
  )
}

function shouldDropLikelyFooterNumberOrDate(el, slideHeight) {
  if (!el || el.type !== 'text') return false
  if (!Number.isFinite(slideHeight) || slideHeight <= 0) return false
  if (!Number.isFinite(el.top) || !Number.isFinite(el.height)) return false

  const plain = stripHtmlToPlainText(el.content)
  if (!plain) return false

  const nearBottom = el.top > slideHeight * 0.78
  const smallBox = el.height < slideHeight * 0.2
  if (!nearBottom || !smallBox) return false

  const looksLikeNumber = /^[0-9]{1,3}$/.test(plain)
  const looksLikeDate = /^(\d{4}[\-/]\d{1,2}[\-/]\d{1,2})$/.test(plain) || /(年\d{1,2}月\d{1,2}日)$/.test(plain)
  const looksLikeKeyword = /页码|日期/.test(plain)

  return looksLikeNumber || looksLikeDate || looksLikeKeyword
}

function filterElementsTree(elements, slideHeight) {
  const out = []
  for (const el of elements || []) {
    if (!el) continue
    if (isLikelyHeaderFooterName(el.name)) continue
    if (shouldDropLikelyFooterNumberOrDate(el, slideHeight)) continue

    if (el.type === 'text' && isPlaceholderPromptText(stripHtmlToPlainText(el.content))) continue

    if (Array.isArray(el.elements)) {
      const nextChildren = filterElementsTree(el.elements, slideHeight)
      if (!nextChildren.length) continue
      out.push({
        ...el,
        elements: nextChildren,
      })
      continue
    }
    out.push(el)
  }
  return out
}

function renumberSiblingOrder(elements) {
  if (!Array.isArray(elements)) return elements

  for (let i = 0; i < elements.length; i++) {
    const el = elements[i]
    if (!el || typeof el !== 'object') continue
    el.order = i
    if (Array.isArray(el.elements)) renumberSiblingOrder(el.elements)
  }

  return elements
}

function filterSlideOutput(slide, slideHeight) {
  const elements = filterElementsTree(slide.elements, slideHeight)
  const layoutElements = filterElementsTree(slide.layoutElements, slideHeight)
  renumberSiblingOrder(elements)
  renumberSiblingOrder(layoutElements)

  return {
    ...slide,
    elements,
    layoutElements,
  }
}

function pushTrace(warpObj, step, data) {
  if (!warpObj || !Array.isArray(warpObj.trace)) return
  if (data === undefined) warpObj.trace.push({ step })
  else warpObj.trace.push({ step, data })
}

export async function parse(file, options = {}) {
  const slides = []
  
  const zip = await JSZip.loadAsync(file)

  const filesInfo = await getContentTypes(zip)
  const { width, height, defaultTextStyle, headerFooter } = await getSlideInfo(zip)
  const { themeContent, themeColors } = await getTheme(zip)

  const orderedSlides = await getSlidesInPresentationOrder(zip)
  const slideFiles = orderedSlides.length ? orderedSlides : filesInfo.slides

  for (let i = 0; i < slideFiles.length; i++) {
    const filename = slideFiles[i]
    const slideNo = i + 1
    const singleSlide = await processSingleSlide(zip, filename, themeContent, defaultTextStyle, headerFooter, slideNo, options)
    slides.push(filterSlideOutput(singleSlide, height))
  }

  return {
    slides,
    themeColors,
    size: {
      width,
      height,
    },
  }
}

async function getContentTypes(zip) {
  const ContentTypesJson = await readXmlFile(zip, '[Content_Types].xml')
  const subObj = ContentTypesJson['Types']['Override']
  let slidesLocArray = []
  let slideLayoutsLocArray = []

  for (const item of subObj) {
    switch (item['attrs']['ContentType']) {
      case 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml':
        slidesLocArray.push(item['attrs']['PartName'].substr(1))
        break
      case 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml':
        slideLayoutsLocArray.push(item['attrs']['PartName'].substr(1))
        break
      default:
    }
  }
  
  const sortSlideXml = (p1, p2) => {
    const n1 = +/(\d+)\.xml/.exec(p1)[1]
    const n2 = +/(\d+)\.xml/.exec(p2)[1]
    return n1 - n2
  }
  slidesLocArray = slidesLocArray.sort(sortSlideXml)
  slideLayoutsLocArray = slideLayoutsLocArray.sort(sortSlideXml)
  
  return {
    slides: slidesLocArray,
    slideLayouts: slideLayoutsLocArray,
  }
}

async function getSlideInfo(zip) {
  const content = await readXmlFile(zip, 'ppt/presentation.xml')
  const sldSzAttrs = content['p:presentation']['p:sldSz']['attrs']
  const defaultTextStyle = content['p:presentation']['p:defaultTextStyle']

  const normalizeOn = (v) => {
    if (v === undefined || v === null) return undefined
    const s = String(v).toLowerCase()
    return s === '1' || s === 'true' || s === 'on'
  }

  const defaultOn = (v) => (v === undefined ? true : v)

  const hfAttrs = getTextByPathList(content, ['p:presentation', 'p:hf', 'attrs'])
  const headerFooter = {
    dt: defaultOn(normalizeOn(getTextByPathList(hfAttrs, ['dt']))),
    ftr: defaultOn(normalizeOn(getTextByPathList(hfAttrs, ['ftr']))),
    hdr: defaultOn(normalizeOn(getTextByPathList(hfAttrs, ['hdr']))),
    sldNum: defaultOn(normalizeOn(getTextByPathList(hfAttrs, ['sldNum']))),
  }
  return {
    width: parseInt(sldSzAttrs['cx']) * RATIO_EMUs_Points,
    height: parseInt(sldSzAttrs['cy']) * RATIO_EMUs_Points,
    defaultTextStyle,
    headerFooter,
  }
}

async function getTheme(zip) {
  const preResContent = await readXmlFile(zip, 'ppt/_rels/presentation.xml.rels')
  const relationshipArray = preResContent['Relationships']['Relationship']
  let themeURI

  if (relationshipArray.constructor === Array) {
    for (const relationshipItem of relationshipArray) {
      if (relationshipItem['attrs']['Type'] === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme') {
        themeURI = relationshipItem['attrs']['Target']
        break
      }
    }
  } 
  else if (relationshipArray['attrs']['Type'] === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme') {
    themeURI = relationshipArray['attrs']['Target']
  }

  const themeContent = await readXmlFile(zip, 'ppt/' + themeURI)

  const themeColors = []
  const clrScheme = getTextByPathList(themeContent, ['a:theme', 'a:themeElements', 'a:clrScheme'])
  if (clrScheme) {
    const keys = Object.keys(clrScheme)
    for (const key of keys) {
      if (!key.startsWith('a:')) continue
      const refNode = clrScheme[key]
      const srgb = getTextByPathList(refNode, ['a:srgbClr', 'attrs', 'val'])
      const sys = getTextByPathList(refNode, ['a:sysClr', 'attrs', 'lastClr'])
      const val = srgb || sys
      if (val) themeColors.push('#' + val)
    }
  }

  return { themeContent, themeColors }
}

async function getSlidesInPresentationOrder(zip) {
  const presentation = await readXmlFile(zip, 'ppt/presentation.xml')
  const presentationRels = await readXmlFile(zip, 'ppt/_rels/presentation.xml.rels')
  if (!presentation || !presentationRels) return []

  const sldIdsRaw = getTextByPathList(presentation, ['p:presentation', 'p:sldIdLst', 'p:sldId'])
  const sldIds = Array.isArray(sldIdsRaw) ? sldIdsRaw : (sldIdsRaw ? [sldIdsRaw] : [])

  const relsRaw = getTextByPathList(presentationRels, ['Relationships', 'Relationship'])
  const rels = Array.isArray(relsRaw) ? relsRaw : (relsRaw ? [relsRaw] : [])
  const slideRelMap = new Map()
  for (const r of rels) {
    const type = getTextByPathList(r, ['attrs', 'Type'])
    if (type !== 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide') continue
    const id = getTextByPathList(r, ['attrs', 'Id'])
    const target = getTextByPathList(r, ['attrs', 'Target'])
    if (!id || !target) continue
    slideRelMap.set(String(id), String(target))
  }

  const out = []
  for (const s of sldIds) {
    const rId = getTextByPathList(s, ['attrs', 'r:id'])
    if (!rId) continue
    const target = slideRelMap.get(String(rId))
    if (!target) continue

    if (target.startsWith('ppt/')) out.push(target)
    else if (target.startsWith('../')) out.push(target.replace('../', 'ppt/'))
    else out.push('ppt/' + target.replace(/^\//, ''))
  }
  return out
}

async function processSingleSlide(zip, sldFileName, themeContent, defaultTextStyle, headerFooter, slideNo, options = {}) {
  const resName = sldFileName.replace('slides/slide', 'slides/_rels/slide') + '.rels'
  const resContent = await readXmlFile(zip, resName)
  let relationshipArray = resContent['Relationships']['Relationship']
  if (relationshipArray.constructor !== Array) relationshipArray = [relationshipArray]
  
  let noteFilename = ''
  let layoutFilename = ''
  let masterFilename = ''
  let themeFilename = ''
  const diagramDrawingTargets = []
  const slideResObj = {}
  const layoutResObj = {}
  const masterResObj = {}
  const themeResObj = {}
  const diagramResObj = {}

  for (const relationshipArrayItem of relationshipArray) {
    switch (relationshipArrayItem['attrs']['Type']) {
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout':
        layoutFilename = relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
        slideResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
          target: relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
        }
        break
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide':
        noteFilename = relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
        slideResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
          target: relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
        }
        break
      case 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing':
        diagramDrawingTargets.push(relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/'))
        slideResObj[relationshipArrayItem['attrs']['Id']] = {
          type: 'diagramDrawing',
          target: relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
        }
        break
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image':
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart':
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink':
      default:
        slideResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
          target: relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/'),
        }
    }
  }

  const trace = options && options.trace ? [] : null
  const traceEnabled = !!trace

  if (traceEnabled) {
    trace.push({
      step: 'slide/rels',
      data: {
        slideNo,
        sldFileName,
        slideRels: resName,
        noteFilename,
        layoutFilename,
      },
    })
  }
  
  const slideNotesContent = await readXmlFile(zip, noteFilename)
  const note = getNote(slideNotesContent)

  const slideLayoutContent = await readXmlFile(zip, layoutFilename)
  const slideLayoutTables = await indexNodes(slideLayoutContent)
  const slideLayoutResFilename = layoutFilename.replace('slideLayouts/slideLayout', 'slideLayouts/_rels/slideLayout') + '.rels'
  const slideLayoutResContent = await readXmlFile(zip, slideLayoutResFilename)
  relationshipArray = slideLayoutResContent['Relationships']['Relationship']
  if (relationshipArray.constructor !== Array) relationshipArray = [relationshipArray]

  for (const relationshipArrayItem of relationshipArray) {
    switch (relationshipArrayItem['attrs']['Type']) {
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster':
        masterFilename = relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
        break
      default:
        layoutResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
          target: relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/'),
        }
    }
  }

  if (traceEnabled) {
    trace.push({
      step: 'slideLayout/rels',
      data: {
        slideLayout: layoutFilename,
        slideLayoutRels: slideLayoutResFilename,
        slideMaster: masterFilename,
      },
    })
  }

  const slideMasterContent = await readXmlFile(zip, masterFilename)
  const slideMasterTextStyles = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:txStyles'])
  const slideMasterTables = indexNodes(slideMasterContent)
  const slideMasterResFilename = masterFilename.replace('slideMasters/slideMaster', 'slideMasters/_rels/slideMaster') + '.rels'
  const slideMasterResContent = await readXmlFile(zip, slideMasterResFilename)
  relationshipArray = slideMasterResContent['Relationships']['Relationship']
  if (relationshipArray.constructor !== Array) relationshipArray = [relationshipArray]

  for (const relationshipArrayItem of relationshipArray) {
    switch (relationshipArrayItem['attrs']['Type']) {
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme':
        themeFilename = relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
        break
      default:
        masterResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
          target: relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/'),
        }
    }
  }

  if (traceEnabled) {
    trace.push({
      step: 'slideMaster/rels',
      data: {
        slideMaster: masterFilename,
        slideMasterRels: slideMasterResFilename,
        themeFilename,
      },
    })
  }

  let slideThemeContent = themeContent
  if (themeFilename) {
    const loadedThemeContent = await readXmlFile(zip, themeFilename)
    if (loadedThemeContent) slideThemeContent = loadedThemeContent

    const themeName = themeFilename.split('/').pop()
    const themeResFileName = themeFilename.replace(themeName, '_rels/' + themeName) + '.rels'
    const themeResContent = await readXmlFile(zip, themeResFileName)
    if (themeResContent) {
      relationshipArray = themeResContent['Relationships']['Relationship']
      if (relationshipArray) {
        if (relationshipArray.constructor !== Array) relationshipArray = [relationshipArray]
        for (const relationshipArrayItem of relationshipArray) {
          themeResObj[relationshipArrayItem['attrs']['Id']] = {
            'type': relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            'target': relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
          }
        }
      }
    }
  }

  const diagramDrawingContents = {}
  const diagramResObjByTarget = {}
  if (diagramDrawingTargets.length) {
    for (const diagramFilename of diagramDrawingTargets) {
      const diagName = diagramFilename.split('/').pop()
      const diagramResFileName = diagramFilename.replace(diagName, '_rels/' + diagName) + '.rels'

      let drawingContent = await readXmlFile(zip, diagramFilename)
      if (drawingContent) {
        const drawingContentStr = JSON.stringify(drawingContent).replace(/dsp:/g, 'p:')
        drawingContent = JSON.parse(drawingContentStr)
      }
      diagramDrawingContents[diagramFilename] = drawingContent

      const currentDiagramResObj = {}
      const digramResContent = await readXmlFile(zip, diagramResFileName)
      if (digramResContent) {
        relationshipArray = digramResContent['Relationships']['Relationship']
        if (relationshipArray.constructor !== Array) relationshipArray = [relationshipArray]
        for (const relationshipArrayItem of relationshipArray) {
          currentDiagramResObj[relationshipArrayItem['attrs']['Id']] = {
            type: relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            target: relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/'),
          }
        }
      }
      diagramResObjByTarget[diagramFilename] = currentDiagramResObj
    }
  }

  const digramFileContent = diagramDrawingTargets.length ? diagramDrawingContents[diagramDrawingTargets[0]] : null
  if (diagramDrawingTargets.length) {
    const firstTarget = diagramDrawingTargets[0]
    const firstResObj = diagramResObjByTarget[firstTarget]
    if (firstResObj) {
      for (const k in firstResObj) diagramResObj[k] = firstResObj[k]
    }
  }

  const tableStyles = await readXmlFile(zip, 'ppt/tableStyles.xml')

  const slideContent = await readXmlFile(zip, sldFileName)
  const nodes = slideContent['p:sld']['p:cSld']['p:spTree']

  const slideAttrs = getTextByPathList(slideContent, ['p:sld', 'attrs'])
  const showPh = getTextByPathList(slideAttrs, ['showPh'])

  const warpObj = {
    zip,
    slideLayoutContent,
    slideLayoutTables,
    slideMasterContent,
    slideMasterTables,
    slideContent,
    tableStyles,
    slideResObj,
    slideMasterTextStyles,
    layoutResObj,
    masterResObj,
    themeContent: slideThemeContent,
    themeResObj,
    digramFileContent,
    diagramResObj,
    diagramDrawingTargets,
    diagramDrawingContents,
    diagramResObjByTarget,
    diagramDrawingCursor: 0,
    defaultTextStyle,
    headerFooter,
    slideNo,
    trace,
  }
  const layoutElements = await getLayoutElements(warpObj)
  const fill = await getSlideBackgroundFill(warpObj)

  const elements = []
  for (const nodeKey in nodes) {
    if (nodes[nodeKey].constructor !== Array) nodes[nodeKey] = [nodes[nodeKey]]
    for (const node of nodes[nodeKey]) {
      if (showPh === '0') {
        const ph =
          getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph']) ||
          getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'p:ph']) ||
          getTextByPathList(node, ['p:nvGraphicFramePr', 'p:nvPr', 'p:ph'])
        if (ph) continue
      }

      const phType = getPlaceholderType(node)
      if (isHeaderFooterPlaceholderType(phType)) continue

      const ret = await processNodesInSlide(nodeKey, node, nodes, warpObj, 'slide')
      if (ret) elements.push(ret)
    }
  }

  sortElementsByOrder(elements)
  sortElementsByOrder(layoutElements)

  let transitionNode = findTransitionNode(slideContent, 'p:sld')
  if (!transitionNode) transitionNode = findTransitionNode(slideLayoutContent, 'p:sldLayout')
  if (!transitionNode) transitionNode = findTransitionNode(slideMasterContent, 'p:sldMaster')

  const transition = parseTransition(transitionNode)

  const out = {
    fill,
    elements,
    layoutElements,
    note,
    transition,
  }

  if (traceEnabled && trace.length) out.trace = trace

  return out
}

function getNote(noteContent) {
  let text = ''
  let spNodes = getTextByPathList(noteContent, ['p:notes', 'p:cSld', 'p:spTree', 'p:sp'])
  if (!spNodes) return ''

  if (spNodes.constructor !== Array) spNodes = [spNodes]
  for (const spNode of spNodes) {
    let rNodes = getTextByPathList(spNode, ['p:txBody', 'a:p', 'a:r'])
    if (!rNodes) continue

    if (rNodes.constructor !== Array) rNodes = [rNodes]
    for (const rNode of rNodes) {
      const t = getTextByPathList(rNode, ['a:t'])
      if (t && typeof t === 'string') text += t
    }
  }
  return text
}

async function getLayoutElements(warpObj) {
  const elements = []
  const slideLayoutContent = warpObj['slideLayoutContent']
  const slideMasterContent = warpObj['slideMasterContent']
  const slideContent = warpObj['slideContent']
  const nodesSldLayout = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:cSld', 'p:spTree'])
  const nodesSldMaster = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:cSld', 'p:spTree'])

  const placeholderOverrides = new Set()
  const slideSpTree = getTextByPathList(slideContent, ['p:sld', 'p:cSld', 'p:spTree'])

  const slideAttrs = getTextByPathList(slideContent, ['p:sld', 'attrs'])
  const slideShowMasterSp = getTextByPathList(slideAttrs, ['showMasterSp'])
  const slideShowMasterPh = getTextByPathList(slideAttrs, ['showMasterPh'])
  const slideShowPh = getTextByPathList(slideAttrs, ['showPh'])

  pushTrace(warpObj, 'layout/start', {
    slideNo: warpObj && warpObj.slideNo,
    slideShowPh,
    slideShowMasterSp,
    slideShowMasterPh,
  })

  const addOverrideKeys = (type, idx) => {
    const t = type || ''
    const i = idx || ''
    if (!t && !i) return
    placeholderOverrides.add(`${t}|${i}`)
    if (t) placeholderOverrides.add(`${t}|`)
    if (i) placeholderOverrides.add(`|${i}`)
  }

  const hasTxBodyText = (n) => {
    const txBody = getTextByPathList(n, ['p:txBody'])
    if (!txBody) return false

    const pNodes = getTextByPathList(txBody, ['a:p'])
    const pList = Array.isArray(pNodes) ? pNodes : (pNodes ? [pNodes] : [])
    for (const p of pList) {
      const rs = getTextByPathList(p, ['a:r'])
      const rList = Array.isArray(rs) ? rs : (rs ? [rs] : [])
      for (const r of rList) {
        const t = getTextByPathList(r, ['a:t'])
        if (typeof t === 'string' && t.trim() !== '') return true
      }

      const flds = getTextByPathList(p, ['a:fld'])
      const fList = Array.isArray(flds) ? flds : (flds ? [flds] : [])
      for (const f of fList) {
        const t = getTextByPathList(f, ['a:t'])
        if (typeof t === 'string' && t.trim() !== '') return true
      }
    }

    return false
  }

  const isMasterPlaceholderRenderable = (node, parts) => {
    if (!parts || !parts.type) return false

    const t = String(parts.type)
    const allowTypes = new Set(['ftr'])
    if (!allowTypes.has(t)) return false

    const normalizeOn = (v) => {
      if (v === undefined || v === null) return false
      const s = String(v).toLowerCase()
      return s === '1' || s === 'true' || s === 'on'
    }

    const slideHdrFtrAttrs = getTextByPathList(warpObj, ['slideContent', 'p:sld', 'p:hdrFtr', 'attrs'])
    const slideFlag = getTextByPathList(slideHdrFtrAttrs, [t])
    const isEnabledBySlide = slideFlag !== undefined ? normalizeOn(slideFlag) : undefined
    const isEnabledByPresentation = warpObj && warpObj.headerFooter && warpObj.headerFooter[t] !== undefined ? !!warpObj.headerFooter[t] : undefined
    const isEnabled = isEnabledBySlide !== undefined ? isEnabledBySlide : (isEnabledByPresentation !== undefined ? isEnabledByPresentation : true)
    if (!isEnabled) return false

    if (t === 'ftr') return hasTxBodyText(node)

    return false
  }

  const isPlaceholderNode = (node) => {
    const ph =
      getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph']) ||
      getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'p:ph']) ||
      getTextByPathList(node, ['p:nvGraphicFramePr', 'p:nvPr', 'p:ph'])
    return !!ph
  }

  const extractPlaceholderKeyParts = node => {
    const ph =
      getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph']) ||
      getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'p:ph']) ||
      getTextByPathList(node, ['p:nvGraphicFramePr', 'p:nvPr', 'p:ph'])
    if (!ph) return null
    const type = getTextByPathList(ph, ['attrs', 'type'])
    const idx = getTextByPathList(ph, ['attrs', 'idx'])
    if (!type && !idx) return null
    return { type, idx }
  }

  const walkAndCollectOverrides = (nodeKey, node) => {
    if (!node) return

    const parts = extractPlaceholderKeyParts(node)
    if (parts) addOverrideKeys(parts.type, parts.idx)

    if (nodeKey === 'p:grpSp') {
      for (const k in node) {
        if (k === 'p:nvGrpSpPr' || k === 'p:grpSpPr') continue
        const v = node[k]
        if (Array.isArray(v)) {
          for (const item of v) walkAndCollectOverrides(k, item)
        }
        else {
          walkAndCollectOverrides(k, v)
        }
      }
      return
    }

    if (nodeKey === 'mc:AlternateContent') {
      const fallback = getTextByPathList(node, ['mc:Fallback'])
      const fallbackGroup = fallback && (fallback['p:grpSp'] || fallback)
      if (fallbackGroup) {
        if (fallbackGroup['p:grpSpPr']) walkAndCollectOverrides('p:grpSp', fallbackGroup)
        else {
          for (const k in fallbackGroup) {
            const v = fallbackGroup[k]
            if (Array.isArray(v)) {
              for (const item of v) walkAndCollectOverrides(k, item)
            }
            else walkAndCollectOverrides(k, v)
          }
        }
      }
    }
  }

  const addOverridesFromSpTree = (spTreeNode) => {
    if (!spTreeNode) return
    for (const nodeKey in spTreeNode) {
      const v = spTreeNode[nodeKey]
      if (Array.isArray(v)) {
        for (const item of v) walkAndCollectOverrides(nodeKey, item)
      }
      else {
        walkAndCollectOverrides(nodeKey, v)
      }
    }
  }

  const isOverriddenByPlaceholderSet = (type, idx) => {
    const t = type || ''
    const i = idx || ''
    if (placeholderOverrides.has(`${t}|${i}`)) return true
    if (t && placeholderOverrides.has(`${t}|`)) return true
    if (i && placeholderOverrides.has(`|${i}`)) return true
    return false
  }

  const pruneLayoutElement = (el) => {
    if (!el) return null
    if (isOverriddenByPlaceholderSet(el.placeholderType, el.placeholderIdx)) return null

    if (Array.isArray(el.elements)) {
      const nextChildren = []
      for (const child of el.elements) {
        const next = pruneLayoutElement(child)
        if (next) nextChildren.push(next)
      }
      if (!nextChildren.length) return null
      return {
        ...el,
        elements: nextChildren,
      }
    }
    return el
  }

  addOverridesFromSpTree(slideSpTree)
  pushTrace(warpObj, 'layout/overrides/fromSlide', { size: placeholderOverrides.size })

  const showMasterSp = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'attrs', 'showMasterSp'])
  const showMasterPh = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'attrs', 'showMasterPh'])

  let layoutKept = 0
  let layoutSkipShowPh = 0
  let layoutSkipOverridden = 0
  let layoutSkipNotRenderable = 0

  if (nodesSldLayout) {
    for (const nodeKey in nodesSldLayout) {
      if (nodesSldLayout[nodeKey].constructor === Array) {
        for (let i = 0; i < nodesSldLayout[nodeKey].length; i++) {
          const currentNode = nodesSldLayout[nodeKey][i]
          if (slideShowPh === '0' && isPlaceholderNode(currentNode)) {
            layoutSkipShowPh += 1
            continue
          }
          const parts = extractPlaceholderKeyParts(currentNode)
          if (parts) {
            const t = parts.type || ''
            const idx = parts.idx || ''
            if (placeholderOverrides.has(`${t}|${idx}`) || (t && placeholderOverrides.has(`${t}|`)) || (idx && placeholderOverrides.has(`|${idx}`))) {
              layoutSkipOverridden += 1
              continue
            }
            if (!isMasterPlaceholderRenderable(currentNode, parts)) {
              layoutSkipNotRenderable += 1
              continue
            }
          }

          const ret = await processNodesInSlide(nodeKey, currentNode, nodesSldLayout, warpObj, 'slideLayoutBg')
          const pruned = pruneLayoutElement(ret)
          if (pruned) {
            elements.push(pruned)
            layoutKept += 1
          }
        }
      } 
      else {
        const currentNode = nodesSldLayout[nodeKey]
        if (slideShowPh === '0' && isPlaceholderNode(currentNode)) {
          layoutSkipShowPh += 1
          continue
        }
        const parts = extractPlaceholderKeyParts(currentNode)
        if (parts) {
          const t = parts.type || ''
          const idx = parts.idx || ''
          if (placeholderOverrides.has(`${t}|${idx}`) || (t && placeholderOverrides.has(`${t}|`)) || (idx && placeholderOverrides.has(`|${idx}`))) {
            layoutSkipOverridden += 1
            continue
          }
          if (!isMasterPlaceholderRenderable(currentNode, parts)) {
            layoutSkipNotRenderable += 1
            continue
          }
        }

        const ret = await processNodesInSlide(nodeKey, currentNode, nodesSldLayout, warpObj, 'slideLayoutBg')
        const pruned = pruneLayoutElement(ret)
        if (pruned) {
          elements.push(pruned)
          layoutKept += 1
        }
      }
    }
  }

  pushTrace(warpObj, 'layout/layoutElements', {
    showMasterSp,
    showMasterPh,
    kept: layoutKept,
    skipShowPh: layoutSkipShowPh,
    skipOverridden: layoutSkipOverridden,
    skipNotRenderable: layoutSkipNotRenderable,
  })

  addOverridesFromSpTree(nodesSldLayout)
  pushTrace(warpObj, 'layout/overrides/fromLayout', { size: placeholderOverrides.size })

  let masterKept = 0
  let masterSkipShowPh = 0
  let masterSkipShowMasterPh = 0
  let masterSkipOverridden = 0
  let masterSkipNotRenderable = 0

  if (nodesSldMaster && showMasterSp !== '0' && slideShowMasterSp !== '0') {
    for (const nodeKey in nodesSldMaster) {
      if (nodesSldMaster[nodeKey].constructor === Array) {
        for (let i = 0; i < nodesSldMaster[nodeKey].length; i++) {
          const currentNode = nodesSldMaster[nodeKey][i]
          if (slideShowPh === '0' && isPlaceholderNode(currentNode)) {
            masterSkipShowPh += 1
            continue
          }
          const parts = extractPlaceholderKeyParts(currentNode)
          if (parts) {
            if (showMasterPh === '0' || slideShowMasterPh === '0') {
              masterSkipShowMasterPh += 1
              continue
            }
            const t = parts.type || ''
            const idx = parts.idx || ''
            if (placeholderOverrides.has(`${t}|${idx}`) || (t && placeholderOverrides.has(`${t}|`)) || (idx && placeholderOverrides.has(`|${idx}`))) {
              masterSkipOverridden += 1
              continue
            }
            if (!isMasterPlaceholderRenderable(currentNode, parts)) {
              masterSkipNotRenderable += 1
              continue
            }
          }

          const ret = await processNodesInSlide(nodeKey, currentNode, nodesSldMaster, warpObj, 'slideMasterBg')
          const pruned = pruneLayoutElement(ret)
          if (pruned) {
            elements.push(pruned)
            masterKept += 1
          }
        }
      } 
      else {
        const currentNode = nodesSldMaster[nodeKey]
        if (slideShowPh === '0' && isPlaceholderNode(currentNode)) {
          masterSkipShowPh += 1
          continue
        }
        const parts = extractPlaceholderKeyParts(currentNode)
        if (parts) {
          if (showMasterPh === '0' || slideShowMasterPh === '0') {
            masterSkipShowMasterPh += 1
            continue
          }
          const t = parts.type || ''
          const idx = parts.idx || ''
          if (placeholderOverrides.has(`${t}|${idx}`) || (t && placeholderOverrides.has(`${t}|`)) || (idx && placeholderOverrides.has(`|${idx}`))) {
            masterSkipOverridden += 1
            continue
          }
          if (!isMasterPlaceholderRenderable(currentNode, parts)) {
            masterSkipNotRenderable += 1
            continue
          }
        }

        const ret = await processNodesInSlide(nodeKey, currentNode, nodesSldMaster, warpObj, 'slideMasterBg')
        const pruned = pruneLayoutElement(ret)
        if (pruned) {
          elements.push(pruned)
          masterKept += 1
        }
      }
    }
  }

  pushTrace(warpObj, 'layout/masterElements', {
    kept: masterKept,
    skipShowPh: masterSkipShowPh,
    skipShowMasterPh: masterSkipShowMasterPh,
    skipOverridden: masterSkipOverridden,
    skipNotRenderable: masterSkipNotRenderable,
  })
  return elements
}

function sortElementsByOrder(elements) {
  return elements.sort((a, b) => {
    const ao = parseInt(a && a.order)
    const bo = parseInt(b && b.order)
    const an = isNaN(ao) ? 0 : ao
    const bn = isNaN(bo) ? 0 : bo
    return an - bn
  })
}

function getGroupXfrmNode(groupNode) {
  return (
    getTextByPathList(groupNode, ['p:grpSpPr', 'a:xfrm']) ||
    getTextByPathList(groupNode, ['p:grpSp', 'p:grpSpPr', 'a:xfrm'])
  )
}

function scaleElementTree(element, ws, hs) {
  if (!element) return element

  const next = { ...element }

  if (typeof next.left === 'number') next.left = numberToFixed(next.left * ws)
  if (typeof next.top === 'number') next.top = numberToFixed(next.top * hs)
  if (typeof next.width === 'number') next.width = numberToFixed(next.width * ws)
  if (typeof next.height === 'number') next.height = numberToFixed(next.height * hs)

  if (typeof next.borderWidth === 'number') next.borderWidth = numberToFixed(next.borderWidth * Math.max(ws, hs))

  if (typeof next.path === 'string') {
    next.path = scaleSvgPathData(next.path, ws, hs)
  }

  if (next.content) {
    next.content = scaleContentFont(next.content, hs)
  }

  if (Array.isArray(next.elements)) {
    next.elements = next.elements.map(child => scaleElementTree(child, ws, hs))
  }

  return next
}

function applyGroupFlipToChildren(children, groupWidth, groupHeight, flipH, flipV) {
  if (!Array.isArray(children) || (!flipH && !flipV)) return children

  const gw = Number(groupWidth)
  const gh = Number(groupHeight)
  if (!Number.isFinite(gw) || !Number.isFinite(gh)) return children

  for (const child of children) {
    if (!child || typeof child !== 'object') continue

    const cw = typeof child.width === 'number' ? child.width : 0
    const ch = typeof child.height === 'number' ? child.height : 0

    if (flipH && typeof child.left === 'number') {
      child.left = numberToFixed(gw - child.left - cw)
    }
    if (flipV && typeof child.top === 'number') {
      child.top = numberToFixed(gh - child.top - ch)
    }

    const hasTextContent = typeof child.content === 'string' && hasValidText(child.content)

    if (child.type === 'text' || hasTextContent) {
      child.isFlipH = false
      child.isFlipV = false
    }
    else if (child.type === 'group' && Array.isArray(child.elements)) {
      applyGroupFlipToChildren(child.elements, child.width, child.height, flipH, flipV)
      child.isFlipH = false
      child.isFlipV = false
    }
    else {
      if (flipH) child.isFlipH = !child.isFlipH
      if (flipV) child.isFlipV = !child.isFlipV
      if ((flipH ? 1 : 0) ^ (flipV ? 1 : 0)) {
        if (typeof child.rotate === 'number') child.rotate = numberToFixed(-child.rotate)
      }
    }
  }

  sortElementsByOrder(children)
  return children
}

function scaleContentFont(html, scale) {
  if (scale === 1 || !html) return html
  return html.replace(/(font-size:\s*)([\d.]+)pt/g, (match, prefix, size) => {
    const newSize = parseFloat(size) * scale
    return `${prefix}${numberToFixed(newSize)}pt`
  })
}

function scaleSvgPathData(d, ws, hs) {
  if (!d) return d
  if (ws === 1 && hs === 1) return d

  const tokens = String(d).match(/[a-zA-Z]|[-+]?(?:\d*\.\d+|\d+)(?:e[-+]?\d+)?/g)
  if (!tokens) return d

  const groupLenByCmd = {
    M: 2,
    L: 2,
    T: 2,
    H: 1,
    V: 1,
    C: 6,
    S: 4,
    Q: 4,
    A: 7,
    Z: 0,
  }

  const out = []
  let cmd = null
  let idxInCmd = 0

  for (const token of tokens) {
    if (/^[a-zA-Z]$/.test(token)) {
      cmd = token
      idxInCmd = 0
      out.push(token)
      continue
    }

    const cmdUpper = cmd ? cmd.toUpperCase() : ''
    const groupLen = groupLenByCmd[cmdUpper]
    if (!groupLen) {
      out.push(token)
      continue
    }

    const pos = idxInCmd % groupLen
    let nextToken = token

    if (cmdUpper === 'A' && (pos === 2 || pos === 3 || pos === 4)) {
      nextToken = token
    }
    else {
      const n = parseFloat(token)
      if (!isNaN(n)) {
        let scaled = n
        if (cmdUpper === 'H') scaled = n * ws
        else if (cmdUpper === 'V') scaled = n * hs
        else if (cmdUpper === 'A') {
          if (pos === 0) scaled = n * ws
          else if (pos === 1) scaled = n * hs
          else if (pos === 5) scaled = n * ws
          else if (pos === 6) scaled = n * hs
        }
        else {
          scaled = (pos % 2 === 0) ? (n * ws) : (n * hs)
        }
        nextToken = String(numberToFixed(scaled))
      }
    }

    out.push(nextToken)
    idxInCmd += 1
  }

  return out.join(' ')
}

function getFirstDefinedXfrmAttr(xfrmNodes, attrName) {
  for (const node of xfrmNodes || []) {
    if (!node) continue
    const v = getTextByPathList(node, ['attrs', attrName])
    if (v !== undefined && v !== null) return v
  }
  return undefined
}

function indexNodes(content) {
  const keys = Object.keys(content)
  const spTreeNode = content[keys[0]]['p:cSld']['p:spTree']
  const idTable = {}
  const idxTable = {}
  const typeTable = {}
  const typeIdxTable = {}

  const extractNvPr = (node) => {
    if (!node) return null
    return node['p:nvSpPr'] || node['p:nvPicPr'] || node['p:nvGraphicFramePr'] || null
  }

  const indexSingleNode = (node) => {
    const nv = extractNvPr(node)
    if (!nv) return

    const id = getTextByPathList(nv, ['p:cNvPr', 'attrs', 'id'])
    const idx = getTextByPathList(nv, ['p:nvPr', 'p:ph', 'attrs', 'idx'])
    const type = getTextByPathList(nv, ['p:nvPr', 'p:ph', 'attrs', 'type'])

    if (id) idTable[id] = node
    if (idx) idxTable[idx] = node
    if (type) typeTable[type] = node
    if (type && idx) typeIdxTable[`${type}|${idx}`] = node
  }

  const walkContainer = (container) => {
    if (!container) return
    for (const key in container) {
      if (key === 'p:nvGrpSpPr' || key === 'p:grpSpPr') continue
      const v = container[key]
      if (Array.isArray(v)) {
        for (const item of v) walkNode(key, item)
      }
      else {
        walkNode(key, v)
      }
    }
  }

  const walkNode = (nodeKey, node) => {
    if (!node) return

    if (nodeKey === 'p:sp' || nodeKey === 'p:pic' || nodeKey === 'p:graphicFrame') {
      indexSingleNode(node)
      return
    }

    if (nodeKey === 'p:grpSp') {
      walkContainer(node)
      return
    }

    if (nodeKey === 'mc:AlternateContent') {
      const fallback = getTextByPathList(node, ['mc:Fallback'])
      const fallbackGroup = fallback && (fallback['p:grpSp'] || fallback)
      if (fallbackGroup) {
        if (fallbackGroup['p:grpSpPr']) walkNode('p:grpSp', fallbackGroup)
        else walkContainer(fallbackGroup)
      }
      
    }
  }

  walkContainer(spTreeNode)

  return { idTable, idxTable, typeTable, typeIdxTable }
}

async function processNodesInSlide(nodeKey, nodeValue, nodes, warpObj, source, groupHierarchy = []) {
  let json

  const phType = getPlaceholderType(nodeValue)
  if (isHeaderFooterPlaceholderType(phType)) return null

  const nodeName = getNodeName(nodeValue)
  if (isLikelyHeaderFooterName(nodeName)) return null

  switch (nodeKey) {
    case 'p:sp': // Shape, Text
      json = await processSpNode(nodeValue, nodes, warpObj, source, groupHierarchy)
      break
    case 'p:cxnSp': // Shape, Text
      json = await processCxnSpNode(nodeValue, nodes, warpObj, source, groupHierarchy)
      break
    case 'p:pic': // Image, Video, Audio
      json = await processPicNode(nodeValue, warpObj, source, groupHierarchy)
      break
    case 'p:graphicFrame': // Chart, Diagram, Table
      json = await processGraphicFrameNode(nodeValue, warpObj, source)
      break
    case 'p:grpSp':
      json = await processGroupSpNode(nodeValue, warpObj, source, groupHierarchy)
      break
    case 'mc:AlternateContent':
      {
        const fallback = getTextByPathList(nodeValue, ['mc:Fallback'])
        const fallbackGroup = fallback && (fallback['p:grpSp'] || fallback)
        if (fallbackGroup && getGroupXfrmNode(fallbackGroup)) {
          json = await processGroupSpNode(fallbackGroup, warpObj, source, groupHierarchy)
          break
        }
        if (getTextByPathList(nodeValue, ['mc:Choice'])) {
          json = await processMathNode(nodeValue, warpObj, source)
        }
      }
      break
    default:
  }

  return json
}

async function processMathNode(node, warpObj, source) {
  const choice = getTextByPathList(node, ['mc:Choice'])
  const fallback = getTextByPathList(node, ['mc:Fallback'])

  const order = node['attrs']['order']
  const xfrmNode = getTextByPathList(choice, ['p:sp', 'p:spPr', 'a:xfrm'])
  const { top, left } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)

  const oMath = findOMath(choice)[0]
  const latex = latexFormart(parseOMath(oMath))

  const blipFill = getTextByPathList(fallback, ['p:sp', 'p:spPr', 'a:blipFill'])
  const picBase64 = await getPicFill(source, blipFill, warpObj)

  let text = ''
  if (getTextByPathList(choice, ['p:sp', 'p:txBody', 'a:p', 'a:r'])) {
    const sp = getTextByPathList(choice, ['p:sp'])
    text = genTextBody(sp['p:txBody'], sp, undefined, undefined, warpObj)
  }

  return {
    type: 'math',
    top,
    left,
    width, 
    height,
    latex,
    picBase64,
    text,
    order,
  }
}

async function processGroupSpNode(node, warpObj, source, parentGroupHierarchy = []) {
  const order = node['attrs']['order']
  const xfrmNode = getGroupXfrmNode(node)
  if (!xfrmNode) return null

  const groupName = getTextByPathList(node, ['p:nvGrpSpPr', 'p:cNvPr', 'attrs', 'name']) || ''
  const groupId = getTextByPathList(node, ['p:nvGrpSpPr', 'p:cNvPr', 'attrs', 'id']) || ''

  const x = parseInt(xfrmNode['a:off']['attrs']['x']) * RATIO_EMUs_Points
  const y = parseInt(xfrmNode['a:off']['attrs']['y']) * RATIO_EMUs_Points
  const cx = parseInt(xfrmNode['a:ext']['attrs']['cx']) * RATIO_EMUs_Points
  const cy = parseInt(xfrmNode['a:ext']['attrs']['cy']) * RATIO_EMUs_Points

  const chOffAttrs = getTextByPathList(xfrmNode, ['a:chOff', 'attrs'])
  const chExtAttrs = getTextByPathList(xfrmNode, ['a:chExt', 'attrs'])

  const chx = (chOffAttrs && chOffAttrs['x'] !== undefined) ? (parseInt(chOffAttrs['x']) * RATIO_EMUs_Points) : 0
  const chy = (chOffAttrs && chOffAttrs['y'] !== undefined) ? (parseInt(chOffAttrs['y']) * RATIO_EMUs_Points) : 0
  const chcx = (chExtAttrs && chExtAttrs['cx'] !== undefined) ? (parseInt(chExtAttrs['cx']) * RATIO_EMUs_Points) : cx
  const chcy = (chExtAttrs && chExtAttrs['cy'] !== undefined) ? (parseInt(chExtAttrs['cy']) * RATIO_EMUs_Points) : cy

  const isFlipV = getTextByPathList(xfrmNode, ['attrs', 'flipV']) === '1'
  const isFlipH = getTextByPathList(xfrmNode, ['attrs', 'flipH']) === '1'

  let rotate = getTextByPathList(xfrmNode, ['attrs', 'rot']) || 0
  if (rotate) rotate = angleToDegrees(rotate)

  const ws = (!chcx || isNaN(chcx) || !cx) ? 1 : (cx / chcx)
  const hs = (!chcy || isNaN(chcy) || !cy) ? 1 : (cy / chcy)


  // 构建当前组合层级（将当前组合添加到父级层级中）
  const currentGroupHierarchy = [...parentGroupHierarchy, node]

  const getGroupLabel = (n) => {
    if (!n) return ''
    const name = getTextByPathList(n, ['p:nvGrpSpPr', 'p:cNvPr', 'attrs', 'name'])
    const id = getTextByPathList(n, ['p:nvGrpSpPr', 'p:cNvPr', 'attrs', 'id'])
    const nodeOrder = getTextByPathList(n, ['attrs', 'order'])
    const parts = []
    if (name) parts.push(String(name))
    if (id) parts.push(`#${id}`)
    if (nodeOrder !== undefined) parts.push(`@${nodeOrder}`)
    return parts.join('')
  }

  pushTrace(warpObj, 'group/start', {
    slideNo: warpObj && warpObj.slideNo,
    source,
    name: groupName,
    id: groupId,
    order,
    hierarchy: currentGroupHierarchy.map(getGroupLabel).filter(Boolean),
    x,
    y,
    cx,
    cy,
    chx,
    chy,
    chcx,
    chcy,
    ws,
    hs,
    isFlipV,
    isFlipH,
    rotate,
  })

  const elements = []
  for (const nodeKey in node) {
    if (node[nodeKey].constructor === Array) {
      for (const item of node[nodeKey]) {
        const ret = await processNodesInSlide(nodeKey, item, node, warpObj, source, currentGroupHierarchy)
        if (ret) elements.push(ret)
      }
    }
    else {
      const ret = await processNodesInSlide(nodeKey, node[nodeKey], node, warpObj, source, currentGroupHierarchy)
      if (ret) elements.push(ret)
    }
  }

  sortElementsByOrder(elements)

  let bboxMinX = Infinity
  let bboxMinY = Infinity
  let bboxMaxX = -Infinity
  let bboxMaxY = -Infinity
  for (const el of elements) {
    if (!el || typeof el.left !== 'number' || typeof el.top !== 'number') continue
    const r = el.left + (typeof el.width === 'number' ? el.width : 0)
    const b = el.top + (typeof el.height === 'number' ? el.height : 0)
    bboxMinX = Math.min(bboxMinX, el.left)
    bboxMinY = Math.min(bboxMinY, el.top)
    bboxMaxX = Math.max(bboxMaxX, r)
    bboxMaxY = Math.max(bboxMaxY, b)
  }

  const hasBBox = Number.isFinite(bboxMinX) && Number.isFinite(bboxMinY) && Number.isFinite(bboxMaxX) && Number.isFinite(bboxMaxY)
  const isLooseGroup = !Number.isFinite(cx) || !Number.isFinite(cy) || !cx || !cy || !Number.isFinite(chcx) || !Number.isFinite(chcy) || !chcx || !chcy

  const bboxW = hasBBox ? (bboxMaxX - bboxMinX) : 0
  const bboxH = hasBBox ? (bboxMaxY - bboxMinY) : 0
  const eps = Math.max(1, Math.min(Number.isFinite(cx) ? cx : 0, Number.isFinite(cy) ? cy : 0) * 0.002)

  const errToSlide = hasBBox ? (Math.abs(bboxW - cx) + Math.abs(bboxH - cy) + Math.abs(bboxMinX - x) + Math.abs(bboxMinY - y)) : Infinity
  const errToChild = hasBBox ? (Math.abs(bboxW - chcx) + Math.abs(bboxH - chcy) + Math.abs(bboxMinX - chx) + Math.abs(bboxMinY - chy)) : Infinity
  const isChildCoordAbsToSlide = !isLooseGroup && hasBBox && (errToSlide + eps * 2) < errToChild

  pushTrace(warpObj, 'group/bbox', {
    slideNo: warpObj && warpObj.slideNo,
    source,
    name: groupName,
    id: groupId,
    order,
    children: elements.length,
    hasBBox,
    bboxMinX,
    bboxMinY,
    bboxMaxX,
    bboxMaxY,
    bboxW,
    bboxH,
    eps,
    isLooseGroup,
    errToSlide,
    errToChild,
    isChildCoordAbsToSlide,
  })

  const baseX = isLooseGroup ? (hasBBox ? bboxMinX : 0) : (isChildCoordAbsToSlide ? x : chx)
  const baseY = isLooseGroup ? (hasBBox ? bboxMinY : 0) : (isChildCoordAbsToSlide ? y : chy)
  const effWs = (isLooseGroup || isChildCoordAbsToSlide) ? 1 : ws
  const effHs = (isLooseGroup || isChildCoordAbsToSlide) ? 1 : hs
  const outLeft = isLooseGroup ? numberToFixed(hasBBox ? bboxMinX : x) : numberToFixed(x)
  const outTop = isLooseGroup ? numberToFixed(hasBBox ? bboxMinY : y) : numberToFixed(y)
  const outWidth = isLooseGroup ? numberToFixed(hasBBox ? (bboxMaxX - bboxMinX) : cx) : numberToFixed(cx)
  const outHeight = isLooseGroup ? numberToFixed(hasBBox ? (bboxMaxY - bboxMinY) : cy) : numberToFixed(cy)

  pushTrace(warpObj, 'group/normalize', {
    slideNo: warpObj && warpObj.slideNo,
    source,
    name: groupName,
    id: groupId,
    order,
    baseX,
    baseY,
    effWs,
    effHs,
    outLeft,
    outTop,
    outWidth,
    outHeight,
  })

  const normalizedChildren = elements.map(element => {
    if (!element || typeof element.left !== 'number' || typeof element.top !== 'number') return element

    const translated = {
      ...element,
      left: element.left - baseX,
      top: element.top - baseY,
    }

    return scaleElementTree(translated, effWs, effHs)
  })

  sortElementsByOrder(normalizedChildren)

  applyGroupFlipToChildren(normalizedChildren, outWidth, outHeight, isFlipH, isFlipV)

  pushTrace(warpObj, 'group/end', {
    slideNo: warpObj && warpObj.slideNo,
    source,
    name: groupName,
    id: groupId,
    order,
    normalizedChildren: normalizedChildren.length,
  })

  return {
    type: 'group',
    top: outTop,
    left: outLeft,
    width: outWidth,
    height: outHeight,
    rotate,
    order,
    isFlipV: false,
    isFlipH: false,
    elements: normalizedChildren,
  }
}

async function processSpNode(node, pNode, warpObj, source, groupHierarchy = []) {
  const name = getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr', 'attrs', 'name'])
  const idx = getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'idx'])
  let type = getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
  const order = getTextByPathList(node, ['attrs', 'order'])

  let slideLayoutSpNode, slideMasterSpNode

  const layoutTables = warpObj['slideLayoutTables']
  const masterTables = warpObj['slideMasterTables']

  if (type && idx) {
    const k = `${type}|${idx}`
    slideLayoutSpNode = (layoutTables && layoutTables.typeIdxTable && layoutTables.typeIdxTable[k]) || (layoutTables && layoutTables.idxTable && layoutTables.idxTable[idx]) || (layoutTables && layoutTables.typeTable && layoutTables.typeTable[type])
    slideMasterSpNode = (masterTables && masterTables.typeIdxTable && masterTables.typeIdxTable[k]) || (masterTables && masterTables.idxTable && masterTables.idxTable[idx]) || (masterTables && masterTables.typeTable && masterTables.typeTable[type])
  }
  else if (idx) {
    slideLayoutSpNode = layoutTables && layoutTables.idxTable ? layoutTables.idxTable[idx] : undefined
    slideMasterSpNode = masterTables && masterTables.idxTable ? masterTables.idxTable[idx] : undefined
  }
  else if (type) {
    slideLayoutSpNode = layoutTables && layoutTables.typeTable ? layoutTables.typeTable[type] : undefined
    slideMasterSpNode = masterTables && masterTables.typeTable ? masterTables.typeTable[type] : undefined
  }

  if (!type) {
    const txBoxVal = getTextByPathList(node, ['p:nvSpPr', 'p:cNvSpPr', 'attrs', 'txBox'])
    if (txBoxVal === '1') type = 'text'
  }
  if (!type) type = getTextByPathList(slideLayoutSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
  if (!type) type = getTextByPathList(slideMasterSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])

  if (!type) {
    if (source === 'diagramBg') type = 'diagram'
    else type = 'obj'
  }

  return await genShape(node, pNode, slideLayoutSpNode, slideMasterSpNode, name, type, order, warpObj, source, groupHierarchy)
}

async function processCxnSpNode(node, pNode, warpObj, source, groupHierarchy = []) {
  const name = node['p:nvCxnSpPr']['p:cNvPr']['attrs']['name']
  const type = (node['p:nvCxnSpPr']['p:nvPr']['p:ph'] === undefined) ? undefined : node['p:nvSpPr']['p:nvPr']['p:ph']['attrs']['type']
  const order = node['attrs']['order']

  return await genShape(node, pNode, undefined, undefined, name, type, order, warpObj, source, groupHierarchy)
}

async function genShape(node, pNode, slideLayoutSpNode, slideMasterSpNode, name, type, order, warpObj, source, groupHierarchy = []) {
  const ph =
    getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph']) ||
    getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'p:ph']) ||
    getTextByPathList(node, ['p:nvGraphicFramePr', 'p:nvPr', 'p:ph'])
  const placeholderType = ph ? (getTextByPathList(ph, ['attrs', 'type']) || '') : ''
  const placeholderIdx = ph ? (getTextByPathList(ph, ['attrs', 'idx']) || '') : ''

  const xfrmList = ['p:spPr', 'a:xfrm']
  const slideXfrmNode = getTextByPathList(node, xfrmList)
  const slideLayoutXfrmNode = getTextByPathList(slideLayoutSpNode, xfrmList)
  const slideMasterXfrmNode = getTextByPathList(slideMasterSpNode, xfrmList)

  const shapType = getTextByPathList(node, ['p:spPr', 'a:prstGeom', 'attrs', 'prst'])
  const custShapType = getTextByPathList(node, ['p:spPr', 'a:custGeom'])

  const { top, left } = getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)
  const { width, height } = getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)

  const xfrmNodes = [slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode]

  const isFlipV = getFirstDefinedXfrmAttr(xfrmNodes, 'flipV') === '1'
  const isFlipH = getFirstDefinedXfrmAttr(xfrmNodes, 'flipH') === '1'

  const rotate = angleToDegrees(getFirstDefinedXfrmAttr(xfrmNodes, 'rot'))

  const txtXframeNode = getTextByPathList(node, ['p:txXfrm'])
  let txtRotate = rotate
  if (txtXframeNode) {
    const txtXframeRot = getTextByPathList(txtXframeNode, ['attrs', 'rot'])
    if (txtXframeRot) txtRotate = rotate + angleToDegrees(txtXframeRot)
  }

  let content = ''
  if (node['p:txBody']) content = genTextBody(node['p:txBody'], node, slideLayoutSpNode, type, warpObj)

  const { borderColor, borderWidth, borderType, strokeDasharray } = getBorder(node, type, warpObj, groupHierarchy)
  let fill = await getShapeFill(node, pNode, undefined, warpObj, source, groupHierarchy) || ''
  if (shapType === 'arc') fill = ''

  let fixedWidth = width
  let fixedHeight = height
  if (shapType === 'line') {
    const minSize = Math.max(1, borderWidth || 0)
    if (!fixedWidth) fixedWidth = minSize
    if (!fixedHeight) fixedHeight = minSize
  }

  let shadow
  const outerShdwNode = getTextByPathList(node, ['p:spPr', 'a:effectLst', 'a:outerShdw'])
  if (outerShdwNode) shadow = getShadow(outerShdwNode, warpObj)

  const vAlign = getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type)
  const isVertical = getTextByPathList(node, ['p:txBody', 'a:bodyPr', 'attrs', 'vert']) === 'eaVert'
  const autoFit = getTextAutoFit(node, slideLayoutSpNode, slideMasterSpNode)

  const data = {
    left,
    top,
    width: fixedWidth,
    height: fixedHeight,
    borderColor,
    borderWidth,
    borderType,
    borderStrokeDasharray: strokeDasharray,
    fill,
    content,
    isFlipV,
    isFlipH,
    rotate,
    vAlign,
    name,
    order,
    placeholderType,
    placeholderIdx,
  }

  if (shadow) data.shadow = shadow
  if (autoFit) data.autoFit = autoFit

  const isHasValidText = data.content && hasValidText(data.content)

  if (custShapType && type !== 'diagram') {
    const w = fixedWidth
    const h = fixedHeight
    const d = getCustomShapePath(custShapType, w, h)
    if (!isHasValidText) data.content = ''

    return {
      ...data,
      type: 'shape',
      shapType: 'custom',
      path: d,
    }
  }

  let shapePath = ''
  if (shapType) shapePath = getShapePath(shapType, fixedWidth, fixedHeight, node)

  if (shapType && (type === 'obj' || !type || shapType !== 'rect')) {
    if (!isHasValidText) data.content = ''
    return {
      ...data,
      type: 'shape',
      shapType,
      path: shapePath,
    }
  }
  if (shapType && !isHasValidText && (fill || borderWidth)) {
    return {
      ...data,
      type: 'shape',
      content: '',
      shapType,
      path: shapePath,
    }
  }
  return {
    ...data,
    type: 'text',
    isVertical,
    isFlipV: false,
    isFlipH: false,
    rotate: txtRotate,
  }
}

async function processPicNode(node, warpObj, source, groupHierarchy = []) {
  let resObj
  if (source === 'slideMasterBg') resObj = warpObj['masterResObj']
  else if (source === 'slideLayoutBg') resObj = warpObj['layoutResObj']
  else resObj = warpObj['slideResObj']

  const order = node['attrs']['order']

  const ph = getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'p:ph'])
  const placeholderType = ph ? (getTextByPathList(ph, ['attrs', 'type']) || '') : ''
  const placeholderIdx = ph ? (getTextByPathList(ph, ['attrs', 'idx']) || '') : ''
  
  const rid = node['p:blipFill']['a:blip']['attrs']['r:embed']
  const imgName = resObj[rid]['target']
  const imgFileExt = extractFileExtension(imgName).toLowerCase()
  const zip = warpObj['zip']
  const imgArrayBuffer = await zip.file(imgName).async('arraybuffer')

  let xfrmNode = node['p:spPr']['a:xfrm']
  const idx = getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'p:ph', 'attrs', 'idx'])
  const slideLayoutXfrmNode = idx ? getTextByPathList(warpObj['slideLayoutTables'], ['idxTable', idx, 'p:spPr', 'a:xfrm']) : undefined
  const slideMasterXfrmNode = idx ? getTextByPathList(warpObj['slideMasterTables'], ['idxTable', idx, 'p:spPr', 'a:xfrm']) : undefined
  if (!xfrmNode) xfrmNode = slideLayoutXfrmNode || slideMasterXfrmNode

  const mimeType = getMimeType(imgFileExt)
  const { top, left } = getPosition(xfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)
  const { width, height } = getSize(xfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)
  const src = `data:${mimeType};base64,${base64ArrayBuffer(imgArrayBuffer)}`

  const xfrmNodes = [xfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode]
  const isFlipV = getFirstDefinedXfrmAttr(xfrmNodes, 'flipV') === '1'
  const isFlipH = getFirstDefinedXfrmAttr(xfrmNodes, 'flipH') === '1'

  const rotate = angleToDegrees(getFirstDefinedXfrmAttr(xfrmNodes, 'rot'))

  const videoNode = getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:videoFile'])
  let videoRid, videoFile, videoFileExt, videoMimeType, uInt8ArrayVideo, videoBlob
  let isVdeoLink = false

  if (videoNode) {
    videoRid = videoNode['attrs']['r:link']
    videoFile = resObj[videoRid]['target']
    if (isVideoLink(videoFile)) {
      videoFile = escapeHtml(videoFile)
      isVdeoLink = true
    } 
    else {
      videoFileExt = extractFileExtension(videoFile).toLowerCase()
      if (videoFileExt === 'mp4' || videoFileExt === 'webm' || videoFileExt === 'ogg') {
        uInt8ArrayVideo = await zip.file(videoFile).async('arraybuffer')
        videoMimeType = getMimeType(videoFileExt)
        videoBlob = URL.createObjectURL(new Blob([uInt8ArrayVideo], {
          type: videoMimeType
        }))
      }
    }
  }

  const audioNode = getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:audioFile'])
  let audioRid, audioFile, audioFileExt, uInt8ArrayAudio, audioBlob
  if (audioNode) {
    audioRid = audioNode['attrs']['r:link']
    audioFile = resObj[audioRid]['target']
    audioFileExt = extractFileExtension(audioFile).toLowerCase()
    if (audioFileExt === 'mp3' || audioFileExt === 'wav' || audioFileExt === 'ogg') {
      uInt8ArrayAudio = await zip.file(audioFile).async('arraybuffer')
      audioBlob = URL.createObjectURL(new Blob([uInt8ArrayAudio]))
    }
  }

  if (videoNode && !isVdeoLink) {
    return {
      type: 'video',
      top,
      left,
      width, 
      height,
      rotate,
      blob: videoBlob,
      order,
      placeholderType,
      placeholderIdx,
    }
  } 
  if (videoNode && isVdeoLink) {
    return {
      type: 'video',
      top,
      left,
      width, 
      height,
      rotate,
      src: videoFile,
      order,
      placeholderType,
      placeholderIdx,
    }
  }
  if (audioNode) {
    return {
      type: 'audio',
      top,
      left,
      width, 
      height,
      rotate,
      blob: audioBlob,
      order,
      placeholderType,
      placeholderIdx,
    }
  }

  let rect
  const srcRectAttrs = getTextByPathList(node, ['p:blipFill', 'a:srcRect', 'attrs'])
  if (srcRectAttrs && (srcRectAttrs.t || srcRectAttrs.b || srcRectAttrs.l || srcRectAttrs.r)) {
    rect = {}
    if (srcRectAttrs.t) rect.t = srcRectAttrs.t / 1000
    if (srcRectAttrs.b) rect.b = srcRectAttrs.b / 1000
    if (srcRectAttrs.l) rect.l = srcRectAttrs.l / 1000
    if (srcRectAttrs.r) rect.r = srcRectAttrs.r / 1000
  }
  const geom = getTextByPathList(node, ['p:spPr', 'a:prstGeom', 'attrs', 'prst']) || 'rect'

  const { borderColor, borderWidth, borderType, strokeDasharray } = getBorder(node, undefined, warpObj, groupHierarchy)

  const filters = getPicFilters(node['p:blipFill'])

  const imageData = {
    type: 'image',
    top,
    left,
    width,
    height,
    rotate,
    src,
    isFlipV,
    isFlipH,
    order,
    rect,
    geom,
    borderColor,
    borderWidth,
    borderType,
    borderStrokeDasharray: strokeDasharray,
    placeholderType,
    placeholderIdx,
  }

  if (filters) imageData.filters = filters

  return imageData
}

async function processGraphicFrameNode(node, warpObj, source) {
  const ph = getTextByPathList(node, ['p:nvGraphicFramePr', 'p:nvPr', 'p:ph'])
  const placeholderType = ph ? (getTextByPathList(ph, ['attrs', 'type']) || '') : ''
  const placeholderIdx = ph ? (getTextByPathList(ph, ['attrs', 'idx']) || '') : ''

  const graphicTypeUri = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'attrs', 'uri'])
  
  let result
  switch (graphicTypeUri) {
    case 'http://schemas.openxmlformats.org/drawingml/2006/table':
      result = await genTable(node, warpObj)
      break
    case 'http://schemas.openxmlformats.org/drawingml/2006/chart':
      result = await genChart(node, warpObj)
      break
    case 'http://schemas.openxmlformats.org/drawingml/2006/diagram':
      result = await genDiagram(node, warpObj)
      break
    case 'http://schemas.openxmlformats.org/presentationml/2006/ole':
      let oleObjNode = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'mc:AlternateContent', 'mc:Fallback', 'p:oleObj'])
      if (!oleObjNode) oleObjNode = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'p:oleObj'])
      if (oleObjNode) result = await processGroupSpNode(oleObjNode, warpObj, source)
      break
    default:
  }

  if (result && (placeholderType || placeholderIdx)) {
    result.placeholderType = placeholderType
    result.placeholderIdx = placeholderIdx
  }
  return result
}

async function genTable(node, warpObj) {
  const order = node['attrs']['order']
  const tableNode = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'a:tbl'])
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { top, left } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)

  const getTblPr = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'a:tbl', 'a:tblPr'])
  let getColsGrid = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'a:tbl', 'a:tblGrid', 'a:gridCol'])
  if (getColsGrid.constructor !== Array) getColsGrid = [getColsGrid]

  const colWidths = []
  if (getColsGrid) {
    for (const item of getColsGrid) {
      const colWidthParam = getTextByPathList(item, ['attrs', 'w']) || 0
      const colWidth = parseInt(colWidthParam) * RATIO_EMUs_Points
      colWidths.push(colWidth)
    }
  }

  const firstRowAttr = getTblPr['attrs'] ? getTblPr['attrs']['firstRow'] : undefined
  const firstColAttr = getTblPr['attrs'] ? getTblPr['attrs']['firstCol'] : undefined
  const lastRowAttr = getTblPr['attrs'] ? getTblPr['attrs']['lastRow'] : undefined
  const lastColAttr = getTblPr['attrs'] ? getTblPr['attrs']['lastCol'] : undefined
  const bandRowAttr = getTblPr['attrs'] ? getTblPr['attrs']['bandRow'] : undefined
  const bandColAttr = getTblPr['attrs'] ? getTblPr['attrs']['bandCol'] : undefined
  const tblStylAttrObj = {
    isFrstRowAttr: (firstRowAttr && firstRowAttr === '1') ? 1 : 0,
    isFrstColAttr: (firstColAttr && firstColAttr === '1') ? 1 : 0,
    isLstRowAttr: (lastRowAttr && lastRowAttr === '1') ? 1 : 0,
    isLstColAttr: (lastColAttr && lastColAttr === '1') ? 1 : 0,
    isBandRowAttr: (bandRowAttr && bandRowAttr === '1') ? 1 : 0,
    isBandColAttr: (bandColAttr && bandColAttr === '1') ? 1 : 0,
  }

  let thisTblStyle
  const tbleStyleId = getTblPr['a:tableStyleId']
  if (tbleStyleId) {
    const tbleStylList = warpObj['tableStyles']['a:tblStyleLst']['a:tblStyle']
    if (tbleStylList) {
      if (tbleStylList.constructor === Array) {
        for (let k = 0; k < tbleStylList.length; k++) {
          if (tbleStylList[k]['attrs']['styleId'] === tbleStyleId) {
            thisTblStyle = tbleStylList[k]
          }
        }
      } 
      else {
        if (tbleStylList['attrs']['styleId'] === tbleStyleId) {
          thisTblStyle = tbleStylList
        }
      }
    }
  }
  if (thisTblStyle) thisTblStyle['tblStylAttrObj'] = tblStylAttrObj

  let borders = {}
  const tblStyl = getTextByPathList(thisTblStyle, ['a:wholeTbl', 'a:tcStyle'])
  const tblBorderStyl = getTextByPathList(tblStyl, ['a:tcBdr'])
  if (tblBorderStyl) borders = getTableBorders(tblBorderStyl, warpObj)

  let tbl_bgcolor = ''
  let tbl_bgFillschemeClr = getTextByPathList(thisTblStyle, ['a:tblBg', 'a:fillRef'])
  if (tbl_bgFillschemeClr) {
    tbl_bgcolor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj)
  }
  if (tbl_bgFillschemeClr === undefined) {
    tbl_bgFillschemeClr = getTextByPathList(thisTblStyle, ['a:wholeTbl', 'a:tcStyle', 'a:fill', 'a:solidFill'])
    tbl_bgcolor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj)
  }

  let trNodes = tableNode['a:tr']
  if (trNodes.constructor !== Array) trNodes = [trNodes]
  
  const data = []
  const rowHeights = []
  for (let i = 0; i < trNodes.length; i++) {
    const trNode = trNodes[i]
    
    const rowHeightParam = getTextByPathList(trNodes[i], ['attrs', 'h']) || 0
    const rowHeight = parseInt(rowHeightParam) * RATIO_EMUs_Points
    rowHeights.push(rowHeight)

    const {
      fillColor,
      fontColor,
      fontBold,
    } = getTableRowParams(trNodes, i, tblStylAttrObj, thisTblStyle, warpObj)

    const tcNodes = trNode['a:tc']
    const tr = []

    if (tcNodes.constructor === Array) {
      for (let j = 0; j < tcNodes.length; j++) {
        const tcNode = tcNodes[j]
        let a_sorce
        if (j === 0 && tblStylAttrObj['isFrstColAttr'] === 1) {
          a_sorce = 'a:firstCol'
          if (tblStylAttrObj['isLstRowAttr'] === 1 && i === (trNodes.length - 1) && getTextByPathList(thisTblStyle, ['a:seCell'])) {
            a_sorce = 'a:seCell'
          } 
          else if (tblStylAttrObj['isFrstRowAttr'] === 1 && i === 0 &&
            getTextByPathList(thisTblStyle, ['a:neCell'])) {
            a_sorce = 'a:neCell'
          }
        } 
        else if (
          (j > 0 && tblStylAttrObj['isBandColAttr'] === 1) &&
          !(tblStylAttrObj['isFrstColAttr'] === 1 && i === 0) &&
          !(tblStylAttrObj['isLstRowAttr'] === 1 && i === (trNodes.length - 1)) &&
          j !== (tcNodes.length - 1)
        ) {
          if ((j % 2) !== 0) {
            let aBandNode = getTextByPathList(thisTblStyle, ['a:band2V'])
            if (aBandNode === undefined) {
              aBandNode = getTextByPathList(thisTblStyle, ['a:band1V'])
              if (aBandNode) a_sorce = 'a:band2V'
            } 
            else a_sorce = 'a:band2V'
          }
        }
        if (j === (tcNodes.length - 1) && tblStylAttrObj['isLstColAttr'] === 1) {
          a_sorce = 'a:lastCol'
          if (tblStylAttrObj['isLstRowAttr'] === 1 && i === (trNodes.length - 1) && getTextByPathList(thisTblStyle, ['a:swCell'])) {
            a_sorce = 'a:swCell'
          } 
          else if (tblStylAttrObj['isFrstRowAttr'] === 1 && i === 0 && getTextByPathList(thisTblStyle, ['a:nwCell'])) {
            a_sorce = 'a:nwCell'
          }
        }
        const text = genTextBody(tcNode['a:txBody'], tcNode, undefined, undefined, warpObj)
        const cell = await getTableCellParams(tcNode, thisTblStyle, a_sorce, warpObj)
        const td = { text }
        if (cell.rowSpan) td.rowSpan = cell.rowSpan
        if (cell.colSpan) td.colSpan = cell.colSpan
        if (cell.vMerge) td.vMerge = cell.vMerge
        if (cell.hMerge) td.hMerge = cell.hMerge
        if (cell.fontBold || fontBold) td.fontBold = cell.fontBold || fontBold
        if (cell.fontColor || fontColor) td.fontColor = cell.fontColor || fontColor
        if (cell.fillColor || fillColor || tbl_bgcolor) td.fillColor = cell.fillColor || fillColor || tbl_bgcolor
        if (cell.borders) td.borders = cell.borders

        tr.push(td)
      }
    } 
    else {
      let a_sorce
      if (tblStylAttrObj['isFrstColAttr'] === 1 && tblStylAttrObj['isLstRowAttr'] !== 1) {
        a_sorce = 'a:firstCol'
      } 
      else if (tblStylAttrObj['isBandColAttr'] === 1 && tblStylAttrObj['isLstRowAttr'] !== 1) {
        let aBandNode = getTextByPathList(thisTblStyle, ['a:band2V'])
        if (!aBandNode) {
          aBandNode = getTextByPathList(thisTblStyle, ['a:band1V'])
          if (aBandNode) a_sorce = 'a:band2V'
        } 
        else a_sorce = 'a:band2V'
      }
      if (tblStylAttrObj['isLstColAttr'] === 1 && tblStylAttrObj['isLstRowAttr'] !== 1) {
        a_sorce = 'a:lastCol'
      }

      const text = genTextBody(tcNodes['a:txBody'], tcNodes, undefined, undefined, warpObj)
      const cell = await getTableCellParams(tcNodes, thisTblStyle, a_sorce, warpObj)
      const td = { text }
      if (cell.rowSpan) td.rowSpan = cell.rowSpan
      if (cell.colSpan) td.colSpan = cell.colSpan
      if (cell.vMerge) td.vMerge = cell.vMerge
      if (cell.hMerge) td.hMerge = cell.hMerge
      if (cell.fontBold || fontBold) td.fontBold = cell.fontBold || fontBold
      if (cell.fontColor || fontColor) td.fontColor = cell.fontColor || fontColor
      if (cell.fillColor || fillColor || tbl_bgcolor) td.fillColor = cell.fillColor || fillColor || tbl_bgcolor
      if (cell.borders) td.borders = cell.borders

      tr.push(td)
    }
    data.push(tr)
  }

  return {
    type: 'table',
    top,
    left,
    width,
    height,
    data,
    order,
    borders,
    rowHeights,
    colWidths,
  }
}

async function genChart(node, warpObj) {
  const order = node['attrs']['order']
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { top, left } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)

  const rid = node['a:graphic']['a:graphicData']['c:chart']['attrs']['r:id']
  let refName = getTextByPathList(warpObj['slideResObj'], [rid, 'target'])
  if (!refName) refName = getTextByPathList(warpObj['layoutResObj'], [rid, 'target'])
  if (!refName) refName = getTextByPathList(warpObj['masterResObj'], [rid, 'target'])
  if (!refName) return {}

  const content = await readXmlFile(warpObj['zip'], refName)
  const plotArea = getTextByPathList(content, ['c:chartSpace', 'c:chart', 'c:plotArea'])

  const chart = getChartInfo(plotArea, warpObj)

  if (!chart) return {}

  const data = {
    type: 'chart',
    top,
    left,
    width,
    height,
    data: chart.data,
    colors: chart.colors,
    chartType: chart.type,
    order,
  }
  if (chart.marker !== undefined) data.marker = chart.marker
  if (chart.barDir !== undefined) data.barDir = chart.barDir
  if (chart.holeSize !== undefined) data.holeSize = chart.holeSize
  if (chart.grouping !== undefined) data.grouping = chart.grouping
  if (chart.style !== undefined) data.style = chart.style

  return data
}

async function genDiagram(node, warpObj) {
  const order = node['attrs']['order']
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { left, top } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)
  
  const relIdsAttrs =
    getTextByPathList(node, ['a:graphic', 'a:graphicData', 'dgm:relIds', 'attrs']) ||
    getTextByPathList(node, ['a:graphic', 'a:graphicData', 'p:relIds', 'attrs'])

  let diagramDrawingTarget
  if (relIdsAttrs) {
    for (const k of Object.keys(relIdsAttrs)) {
      if (!k.startsWith('r:')) continue
      const rid = relIdsAttrs[k]
      const target = getTextByPathList(warpObj, ['slideResObj', rid, 'target'])
      if (target && typeof target === 'string' && /\/diagrams\/drawing/i.test(target)) {
        diagramDrawingTarget = target
        break
      }
    }
  }

  if (!diagramDrawingTarget) {
    const cursor = warpObj.diagramDrawingCursor || 0
    diagramDrawingTarget = warpObj.diagramDrawingTargets && warpObj.diagramDrawingTargets[cursor]
    warpObj.diagramDrawingCursor = cursor + 1
  }

  const previousDiagramResObj = warpObj.diagramResObj
  const previousDigramFileContent = warpObj.digramFileContent
  if (diagramDrawingTarget) {
    warpObj.diagramResObj = getTextByPathList(warpObj, ['diagramResObjByTarget', diagramDrawingTarget]) || previousDiagramResObj
    warpObj.digramFileContent = getTextByPathList(warpObj, ['diagramDrawingContents', diagramDrawingTarget]) || previousDigramFileContent
  }

  const elements = []
  const spTree = getTextByPathList(warpObj['digramFileContent'], ['p:drawing', 'p:spTree'])
  if (spTree) {
    for (const nodeKey in spTree) {
      if (nodeKey === 'p:nvGrpSpPr' || nodeKey === 'p:grpSpPr') continue
      const spTreeNode = spTree[nodeKey]
      if (Array.isArray(spTreeNode)) {
        for (const item of spTreeNode) {
          const el = await processNodesInSlide(nodeKey, item, spTree, warpObj, 'diagramBg')
          if (el) elements.push(el)
        }
      }
      else {
        const el = await processNodesInSlide(nodeKey, spTreeNode, spTree, warpObj, 'diagramBg')
        if (el) elements.push(el)
      }
    }
  }

  sortElementsByOrder(elements)

  warpObj.diagramResObj = previousDiagramResObj
  warpObj.digramFileContent = previousDigramFileContent

  return {
    type: 'diagram',
    left,
    top,
    width,
    height,
    elements,
    order,
  }
}
