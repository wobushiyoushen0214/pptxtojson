import { getHorizontalAlign } from './align'
import { escapeHtml, getTextByPathList } from './utils'

import {
  getFontType,
  getFontColor,
  getFontSize,
  getFontBold,
  getFontItalic,
  getFontDecoration,
  getFontDecorationLine,
  getFontSpace,
  getFontSubscript,
  getFontShadow,
} from './fontStyle'

export function genTextBody(textBodyNode, spNode, slideLayoutSpNode, type, warpObj) {
  if (!textBodyNode) return ''

  let text = ''

  const pFontStyle = getTextByPathList(spNode, ['p:style', 'a:fontRef'])

  const pNode = textBodyNode['a:p']
  const pNodes = pNode.constructor === Array ? pNode : [pNode]

  let currentListState = null
  const listCounterByKey = new Map()

  for (const pNode of pNodes) {
    let rNode = pNode['a:r']
    let fldNode = pNode['a:fld']
    let brNode = pNode['a:br']
    if (rNode) {
      rNode = (rNode.constructor === Array) ? rNode : [rNode]

      if (fldNode) {
        fldNode = (fldNode.constructor === Array) ? fldNode : [fldNode]
        rNode = rNode.concat(fldNode)
      }
      if (brNode) {
        brNode = (brNode.constructor === Array) ? brNode : [brNode]
        brNode.forEach(item => item.type = 'br')
  
        if (brNode.length > 1) brNode.shift()
        rNode = rNode.concat(brNode)
        rNode.sort((a, b) => {
          if (!a.attrs || !b.attrs) return true
          return a.attrs.order - b.attrs.order
        })
      }
    }

    const align = getHorizontalAlign(pNode, spNode, type, warpObj)

    const listInfo = getListInfo(pNode, textBodyNode, slideLayoutSpNode, type, warpObj)
    if (listInfo) {
      const nextKey = getListKey(listInfo)
      if (!currentListState || currentListState.key !== nextKey) {
        if (currentListState) text += `</${currentListState.tag}>`
        text += `<${listInfo.tag} style="list-style: none; padding-left: 0; margin: 0;">`
        const nextCounter = listInfo.kind === 'autoNum'
          ? (listCounterByKey.has(nextKey) ? listCounterByKey.get(nextKey) : listInfo.startAt)
          : null
        currentListState = {
          key: nextKey,
          tag: listInfo.tag,
          listInfo,
          counter: nextCounter,
        }
      }

      const marker = getListMarker(currentListState)
      if (currentListState.listInfo.kind === 'autoNum') {
        listCounterByKey.set(currentListState.key, currentListState.counter)
      }
      const bulletStyle = getListMarkerStyle(currentListState.listInfo)
      const indent = (listInfo.lvl - 1) * 1.5
      text += `<li style="text-align: ${align}; margin-left: ${indent}em;"><span style="${bulletStyle}">${marker}</span>`
    }
    else {
      if (currentListState) {
        text += `</${currentListState.tag}>`
        currentListState = null
        listCounterByKey.clear()
      }
      text += `<p style="text-align: ${align};">`
    }
    
    if (!rNode) {
      text += genSpanElement(pNode, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, type, warpObj)
    } 
    else {
      let prevStyleInfo = null
      let accumulatedText = ''

      for (const rNodeItem of rNode) {
        const styleInfo = getSpanStyleInfo(rNodeItem, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, type, warpObj)

        if (!prevStyleInfo || prevStyleInfo.styleText !== styleInfo.styleText || prevStyleInfo.hasLink !== styleInfo.hasLink || styleInfo.hasLink) {
          if (accumulatedText) {
            const processedText = accumulatedText.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, '&nbsp;')
            text += `<span style="${prevStyleInfo.styleText}">${processedText}</span>`
            accumulatedText = ''
          }

          if (styleInfo.hasLink) {
            const processedText = styleInfo.text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, '&nbsp;')
            text += `<span style="${styleInfo.styleText}"><a href="${styleInfo.linkURL}" target="_blank">${processedText}</a></span>`
            prevStyleInfo = null
          } 
          else {
            prevStyleInfo = styleInfo
            accumulatedText = styleInfo.text
          }
        } 
        else accumulatedText += styleInfo.text
      }

      if (accumulatedText && prevStyleInfo) {
        const processedText = accumulatedText.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, '&nbsp;')
        text += `<span style="${prevStyleInfo.styleText}">${processedText}</span>`
      }
    }

    if (listInfo) text += '</li>'
    else text += '</p>'
  }
  if (currentListState) text += `</${currentListState.tag}>`
  return text
}

export function getListInfo(node, textBodyNode, slideLayoutSpNode, type, warpObj) {
  const pPrNode = node['a:pPr']
  if (!pPrNode) return null
  if (pPrNode['a:buNone']) return null

  let lvl = 1
  const lvlNode = getTextByPathList(pPrNode, ['attrs', 'lvl'])
  if (lvlNode !== undefined) lvl = parseInt(lvlNode) + 1

  const direct = extractListDefFromPPrLike(pPrNode)
  if (direct) return { ...direct, lvl }

  const fromTextBody = extractListDefFromPPrLike(getTextByPathList(textBodyNode, ['a:lstStyle', `a:lvl${lvl}pPr`]))
  if (fromTextBody) return { ...fromTextBody, lvl }

  const fromLayout = extractListDefFromPPrLike(getTextByPathList(slideLayoutSpNode, ['p:txBody', 'a:lstStyle', `a:lvl${lvl}pPr`]))
  if (fromLayout) return { ...fromLayout, lvl }

  const slideMasterTextStyles = warpObj && warpObj['slideMasterTextStyles']
  if (slideMasterTextStyles) {
    const styleKey = resolveMasterStyleKey(type, slideMasterTextStyles)
    if (styleKey) {
      const fromMaster = extractListDefFromPPrLike(getTextByPathList(slideMasterTextStyles, [styleKey, `a:lvl${lvl}pPr`]))
      if (fromMaster) return { ...fromMaster, lvl }
      const fromMasterLvl1 = extractListDefFromPPrLike(getTextByPathList(slideMasterTextStyles, [styleKey, 'a:lvl1pPr']))
      if (fromMasterLvl1) return { ...fromMasterLvl1, lvl }
    }
  }

  return null
}

function resolveMasterStyleKey(type, slideMasterTextStyles) {
  if (!slideMasterTextStyles) return null
  if (type === 'title' || type === 'ctrTitle') return 'p:titleStyle'
  if (type === 'subTitle') return slideMasterTextStyles['p:titleStyle'] ? 'p:titleStyle' : 'p:bodyStyle'
  if (type === 'body') return 'p:bodyStyle'
  return 'p:otherStyle'
}

function extractListDefFromPPrLike(pPrLikeNode) {
  if (!pPrLikeNode) return null
  if (pPrLikeNode['a:buNone']) return null

  if (pPrLikeNode['a:buChar']) {
    const char = getTextByPathList(pPrLikeNode, ['a:buChar', 'attrs', 'char']) || '•'
    const font = getTextByPathList(pPrLikeNode, ['a:buFont', 'attrs', 'typeface']) || ''
    return {
      kind: 'char',
      tag: 'ul',
      char,
      font,
    }
  }

  if (pPrLikeNode['a:buAutoNum']) {
    const autoNumNode = pPrLikeNode['a:buAutoNum']
    const numType = getTextByPathList(autoNumNode, ['attrs', 'type']) || 'arabicPeriod'
    const startAtRaw = getTextByPathList(autoNumNode, ['attrs', 'startAt'])
    const startAt = startAtRaw ? parseInt(startAtRaw) : 1
    const font = getTextByPathList(pPrLikeNode, ['a:buFont', 'attrs', 'typeface']) || ''
    return {
      kind: 'autoNum',
      tag: 'ol',
      numType,
      startAt: isNaN(startAt) ? 1 : startAt,
      font,
    }
  }

  return null
}

function getListKey(listInfo) {
  if (!listInfo) return ''
  if (listInfo.kind === 'autoNum') return `${listInfo.tag}:${listInfo.kind}:${listInfo.numType}:${listInfo.startAt}:${listInfo.lvl}:${listInfo.font}`
  return `${listInfo.tag}:${listInfo.kind}:${listInfo.char}:${listInfo.lvl}:${listInfo.font}`
}

function getListMarkerStyle(listInfo) {
  let style = 'display: inline-block; min-width: 1.4em; margin-right: 0.4em;'
  if (listInfo.font) style += `font-family: ${listInfo.font};`
  return style
}

function getListMarker(listState) {
  const listInfo = listState.listInfo
  if (listInfo.kind === 'char') return escapeHtml(listInfo.char)

  const n = listState.counter
  listState.counter += 1
  return escapeHtml(formatAutoNumber(n, listInfo.numType))
}

function formatAutoNumber(n, numType) {
  if (/circle/i.test(numType)) return toCircledNumber(n)

  const lowerType = String(numType || '').toLowerCase()
  const isChinese = lowerType.includes('chs') || lowerType.includes('cht')

  const bothParen = numType.includes('ParenBoth')
  let suffix = ''
  if (!bothParen) {
    if (numType.includes('ParenR')) suffix = ')'
    else if (numType.includes('Period')) suffix = '.'
    else if (numType.includes('Comma')) suffix = ','
  }

  let core
  if (numType.includes('alphaLc')) core = toAlpha(n, false)
  else if (numType.includes('alphaUc')) core = toAlpha(n, true)
  else if (numType.includes('romanLc')) core = toRoman(n, false)
  else if (numType.includes('romanUc')) core = toRoman(n, true)
  else if (isChinese) core = toChineseNumber(n, lowerType.includes('cht'), lowerType.includes('db'))
  else core = String(n)

  if (bothParen) return `(${core})`

  const finalSuffix = isChinese && suffix === '.' ? '、' : suffix
  return `${core}${finalSuffix}`
}

function toCircledNumber(n) {
  const num = parseInt(n)
  if (isNaN(num) || num <= 0) return String(n)
  if (num >= 1 && num <= 20) return String.fromCharCode(0x2460 + (num - 1))
  if (num >= 21 && num <= 35) return String.fromCharCode(0x3251 + (num - 21))
  if (num >= 36 && num <= 50) return String.fromCharCode(0x32B1 + (num - 36))
  return String(num)
}

function toAlpha(n, upper) {
  let num = n
  let s = ''
  while (num > 0) {
    num -= 1
    s = String.fromCharCode((num % 26) + 65) + s
    num = Math.floor(num / 26)
  }
  return upper ? s : s.toLowerCase()
}

function toRoman(n, upper) {
  const num = Math.max(1, Math.min(3999, n))
  const map = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ]
  let r = ''
  let v = num
  for (const [value, sym] of map) {
    while (v >= value) {
      r += sym
      v -= value
    }
  }
  return upper ? r : r.toLowerCase()
}

function toChineseNumber(n, traditional, financial) {
  const num = parseInt(n)
  if (isNaN(num) || num <= 0) return String(n)

  const digits = financial
    ? ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九']
    : ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖']

  const units = financial ? ['', '拾', '佰', '仟'] : ['', '十', '百', '千']
  const bigUnits = traditional
    ? ['', '萬', '億', '兆']
    : ['', '万', '亿', '兆']

  const convertUnder10000 = (value) => {
    const v = value % 10000
    if (v === 0) return ''

    const parts = []
    const d3 = Math.floor(v / 1000)
    const d2 = Math.floor((v % 1000) / 100)
    const d1 = Math.floor((v % 100) / 10)
    const d0 = v % 10

    const ds = [d3, d2, d1, d0]
    const us = [units[3], units[2], units[1], units[0]]

    let zeroPending = false
    for (let i = 0; i < ds.length; i++) {
      const d = ds[i]
      const u = us[i]
      const isLast = i === ds.length - 1

      if (d === 0) {
        if (!isLast && parts.length) {
          const rest = ds.slice(i + 1).some(x => x !== 0)
          if (rest) zeroPending = true
        }
        continue
      }

      if (zeroPending) {
        parts.push(digits[0])
        zeroPending = false
      }

      if (u === units[1] && d === 1 && parts.length === 0) {
        parts.push(u)
      }
      else {
        parts.push(digits[d] + u)
      }
    }
    return parts.join('')
  }

  if (num < 10000) return convertUnder10000(num)

  const groups = []
  let remaining = num
  while (remaining > 0) {
    groups.push(remaining % 10000)
    remaining = Math.floor(remaining / 10000)
  }

  let out = ''
  for (let i = groups.length - 1; i >= 0; i--) {
    const g = groups[i]
    if (g === 0) {
      if (out && !out.endsWith(digits[0])) out += digits[0]
      continue
    }

    if (out) {
      const prevWasZero = out.endsWith(digits[0])
      if (!prevWasZero && g < 1000) out += digits[0]
      if (prevWasZero && g >= 1000) out = out.slice(0, -digits[0].length)
    }

    out += convertUnder10000(g) + (bigUnits[i] || '')
  }

  out = out.replace(/零+$/g, '')
  return out || digits[0]
}

export function genSpanElement(node, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, type, warpObj) {
  const { styleText, text, hasLink, linkURL } = getSpanStyleInfo(node, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, type, warpObj)
  const processedText = text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, '&nbsp;')

  if (hasLink) {
    return `<span style="${styleText}"><a href="${linkURL}" target="_blank">${processedText}</a></span>`
  }
  return `<span style="${styleText}">${processedText}</span>`
}

export function getSpanStyleInfo(node, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, type, warpObj) {
  const lstStyle = textBodyNode['a:lstStyle']
  const slideMasterTextStyles = warpObj['slideMasterTextStyles']

  let lvl = 1
  const pPrNode = pNode['a:pPr']
  const lvlNode = getTextByPathList(pPrNode, ['attrs', 'lvl'])
  if (lvlNode !== undefined) lvl = parseInt(lvlNode) + 1

  let text = node['a:t']
  if (typeof text !== 'string') text = getTextByPathList(node, ['a:fld', 'a:t'])
  if (typeof text !== 'string') text = '&nbsp;'

  const fldTypeRaw = getTextByPathList(node, ['attrs', 'type']) || getTextByPathList(node, ['a:fld', 'attrs', 'type'])
  const fldType = typeof fldTypeRaw === 'string' ? fldTypeRaw.toLowerCase() : ''
  if (fldType === 'slidenum' && warpObj && warpObj.slideNo !== undefined && warpObj.slideNo !== null) {
    text = String(warpObj.slideNo)
  }
  else if (fldType.startsWith('datetime')) {
    const trimmed = typeof text === 'string' ? text.replace(/\s+/g, ' ').trim() : ''
    if (!trimmed || trimmed === '日期' || trimmed.toLowerCase() === 'date') {
      const d = new Date()
      const yyyy = String(d.getFullYear())
      const mm = String(d.getMonth() + 1).padStart(2, '0')
      const dd = String(d.getDate()).padStart(2, '0')
      text = `${yyyy}-${mm}-${dd}`
    }
  }

  const plainTrimmed = typeof text === 'string' ? text.replace(/\s+/g, ' ').trim() : ''
  if (type === 'sldNum' && (plainTrimmed === '<#>' || plainTrimmed === '#') && warpObj && warpObj.slideNo !== undefined && warpObj.slideNo !== null) {
    text = String(warpObj.slideNo)
  }
  else if (type === 'dt' && (plainTrimmed === '日期' || plainTrimmed.toLowerCase() === 'date')) {
    const d = new Date()
    const yyyy = String(d.getFullYear())
    const mm = String(d.getMonth() + 1).padStart(2, '0')
    const dd = String(d.getDate()).padStart(2, '0')
    text = `${yyyy}-${mm}-${dd}`
  }

  let styleText = ''
  const fontColor = getFontColor(node, pNode, lstStyle, pFontStyle, lvl, warpObj)
  const fontSize = getFontSize(node, slideLayoutSpNode, type, slideMasterTextStyles, warpObj['defaultTextStyle'], textBodyNode, pNode)
  const fontType = getFontType(node, type, warpObj)
  const fontBold = getFontBold(node)
  const fontItalic = getFontItalic(node)
  const fontDecoration = getFontDecoration(node)
  const fontDecorationLine = getFontDecorationLine(node)
  const fontSpace = getFontSpace(node)
  const shadow = getFontShadow(node, warpObj)
  const subscript = getFontSubscript(node)

  if (fontColor) {
    if (typeof fontColor === 'string') styleText += `color: ${fontColor};`
    else if (fontColor.colors) {
      const { colors, rot } = fontColor
      const stops = colors.map(item => `${item.color} ${item.pos}`).join(', ')
      const gradientStyle = `linear-gradient(${rot + 90}deg, ${stops})`
      styleText += `background: ${gradientStyle}; background-clip: text; color: transparent;`
    }
  }
  if (fontSize) styleText += `font-size: ${fontSize};`
  if (fontType) styleText += `font-family: ${fontType};`
  if (fontBold) styleText += `font-weight: ${fontBold};`
  if (fontItalic) styleText += `font-style: ${fontItalic};`
  if (fontDecoration) styleText += `text-decoration: ${fontDecoration};`
  if (fontDecorationLine) styleText += `text-decoration-line: ${fontDecorationLine};`
  if (fontSpace) styleText += `letter-spacing: ${fontSpace};`
  if (subscript) styleText += `vertical-align: ${subscript};`
  if (shadow) styleText += `text-shadow: ${shadow};`

  const linkID = getTextByPathList(node, ['a:rPr', 'a:hlinkClick', 'attrs', 'r:id'])
  const hasLink = linkID && warpObj['slideResObj'][linkID]

  return {
    styleText,
    text,
    hasLink,
    linkURL: hasLink ? warpObj['slideResObj'][linkID]['target'] : null
  }
}
