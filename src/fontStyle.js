import { getTextByPathList } from './utils'
import { getShadow } from './shadow'
import { getFillType, getGradientFill, getSolidFill } from './fill'

export function getFontType(node, type, warpObj) {
  const text = getTextByPathList(node, ['a:t'])
  const preferEa = typeof text === 'string' && /[\u3040-\u30ff\u3400-\u9fff\uf900-\ufaff]/.test(text)

  let typeface =
    (preferEa ? getTextByPathList(node, ['a:rPr', 'a:ea', 'attrs', 'typeface']) : null) ||
    getTextByPathList(node, ['a:rPr', 'a:latin', 'attrs', 'typeface']) ||
    getTextByPathList(node, ['a:rPr', 'a:ea', 'attrs', 'typeface']) ||
    getTextByPathList(node, ['a:rPr', 'a:cs', 'attrs', 'typeface'])

  if (!typeface) {
    const fontSchemeNode = getTextByPathList(warpObj['themeContent'], ['a:theme', 'a:themeElements', 'a:fontScheme'])

    let fontNode

    if (type === 'title' || type === 'subTitle' || type === 'ctrTitle') {
      fontNode = getTextByPathList(fontSchemeNode, ['a:majorFont'])
    }
    else {
      fontNode = getTextByPathList(fontSchemeNode, ['a:minorFont'])
    }

    typeface =
      (preferEa ? getTextByPathList(fontNode, ['a:ea', 'attrs', 'typeface']) : null) ||
      getTextByPathList(fontNode, ['a:latin', 'attrs', 'typeface']) ||
      getTextByPathList(fontNode, ['a:ea', 'attrs', 'typeface']) ||
      getTextByPathList(fontNode, ['a:cs', 'attrs', 'typeface'])
  }

  return typeface || ''
}

export function getFontColor(node, pNode, lstStyle, pFontStyle, lvl, warpObj) {
  const rPrNode = getTextByPathList(node, ['a:rPr'])
  let filTyp, color
  if (rPrNode) {
    filTyp = getFillType(rPrNode)
    if (filTyp === 'SOLID_FILL') {
      const solidFillNode = rPrNode['a:solidFill']
      color = getSolidFill(solidFillNode, undefined, undefined, warpObj)
    }
    if (filTyp === 'GRADIENT_FILL') {
      const gradientFillNode = rPrNode['a:gradFill']
      const gradient = getGradientFill(gradientFillNode, warpObj)
      return gradient
    }
  }
  if (!color && getTextByPathList(lstStyle, ['a:lvl' + lvl + 'pPr', 'a:defRPr'])) {
    const lstStyledefRPr = getTextByPathList(lstStyle, ['a:lvl' + lvl + 'pPr', 'a:defRPr'])
    filTyp = getFillType(lstStyledefRPr)
    if (filTyp === 'SOLID_FILL') {
      const solidFillNode = lstStyledefRPr['a:solidFill']
      color = getSolidFill(solidFillNode, undefined, undefined, warpObj)
    }
  }
  if (!color) {
    const sPstyle = getTextByPathList(pNode, ['p:style', 'a:fontRef'])
    if (sPstyle) color = getSolidFill(sPstyle, undefined, undefined, warpObj)
    if (!color && pFontStyle) color = getSolidFill(pFontStyle, undefined, undefined, warpObj)
  }
  return color || ''
}

export function getFontSize(node, slideLayoutSpNode, type, slideMasterTextStyles, defaultTextStyle, textBodyNode, pNode) {
  let fontSize

  let lvl = 1
  if (pNode) {
    const lvlNode = getTextByPathList(pNode, ['a:pPr', 'attrs', 'lvl'])
    if (lvlNode !== undefined) lvl = parseInt(lvlNode) + 1
  }

  if (getTextByPathList(node, ['a:rPr', 'attrs', 'sz'])) fontSize = getTextByPathList(node, ['a:rPr', 'attrs', 'sz']) / 100

  if ((isNaN(fontSize) || !fontSize) && pNode) {
    if (getTextByPathList(pNode, ['a:endParaRPr', 'attrs', 'sz'])) {
      fontSize = getTextByPathList(pNode, ['a:endParaRPr', 'attrs', 'sz']) / 100
    }
  }

  if ((isNaN(fontSize) || !fontSize) && textBodyNode) {
    const lstStyle = getTextByPathList(textBodyNode, ['a:lstStyle'])
    if (lstStyle) {
      const sz = getTextByPathList(lstStyle, [`a:lvl${lvl}pPr`, 'a:defRPr', 'attrs', 'sz'])
      if (sz) fontSize = parseInt(sz) / 100
    }
  }

  if ((isNaN(fontSize) || !fontSize) && slideLayoutSpNode) {
    const layoutSz = getTextByPathList(slideLayoutSpNode, ['p:txBody', 'a:lstStyle', `a:lvl${lvl}pPr`, 'a:defRPr', 'attrs', 'sz'])
    if (layoutSz) fontSize = parseInt(layoutSz) / 100
  }

  if ((isNaN(fontSize) || !fontSize) && slideLayoutSpNode) {
    const sz = getTextByPathList(slideLayoutSpNode, ['p:txBody', 'a:lstStyle', 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
    if (sz) fontSize = parseInt(sz) / 100
  }

  if ((isNaN(fontSize) || !fontSize) && pNode) {
    const paraSz = getTextByPathList(pNode, ['a:pPr', 'a:defRPr', 'attrs', 'sz'])
    if (paraSz) fontSize = parseInt(paraSz) / 100
  }

  if (isNaN(fontSize) || !fontSize) {
    if (type === 'dt' || type === 'sldNum') {
      fontSize = 12
    }
    else {
      const lvlKey = `a:lvl${lvl}pPr`

      const tryGetMasterSz = (styleKey) => {
        const masterSz = getTextByPathList(slideMasterTextStyles, [styleKey, lvlKey, 'a:defRPr', 'attrs', 'sz'])
        if (masterSz) return parseInt(masterSz) / 100

        const masterLvl1Sz = getTextByPathList(slideMasterTextStyles, [styleKey, 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
        if (masterLvl1Sz) return parseInt(masterLvl1Sz) / 100

        return null
      }

      let resolved
      if (type === 'title' || type === 'ctrTitle') {
        resolved = tryGetMasterSz('p:titleStyle')
      }
      else if (type === 'subTitle') {
        resolved = tryGetMasterSz('p:titleStyle')
        if (resolved === null) resolved = tryGetMasterSz('p:bodyStyle')
      }
      else if (type === 'body') {
        resolved = tryGetMasterSz('p:bodyStyle')
      }
      else {
        resolved = tryGetMasterSz('p:otherStyle')
      }

      if (resolved !== null) fontSize = resolved
    }
  }

  if (isNaN(fontSize) || !fontSize) {
    const lvlKey = `a:lvl${lvl}pPr`
    const defaultSz =
      getTextByPathList(defaultTextStyle, [lvlKey, 'a:defRPr', 'attrs', 'sz']) ||
      getTextByPathList(defaultTextStyle, ['a:defPPr', 'a:defRPr', 'attrs', 'sz'])
    if (defaultSz) fontSize = parseInt(defaultSz) / 100
  }

  fontSize = (isNaN(fontSize) || !fontSize) ? 18 : fontSize

  return fontSize + 'px'
}

export function getFontBold(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'b']) === '1' ? 'bold' : ''
}

export function getFontItalic(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'i']) === '1' ? 'italic' : ''
}

export function getFontDecoration(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'u']) === 'sng' ? 'underline' : ''
}

export function getFontDecorationLine(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'strike']) === 'sngStrike' ? 'line-through' : ''
}

export function getFontSpace(node) {
  const spc = getTextByPathList(node, ['a:rPr', 'attrs', 'spc'])
  return spc ? (parseInt(spc) / 100 + 'pt') : ''
}

export function getFontSubscript(node) {
  const baseline = getTextByPathList(node, ['a:rPr', 'attrs', 'baseline'])
  if (!baseline) return ''
  return parseInt(baseline) > 0 ? 'super' : 'sub'
}

export function getFontShadow(node, warpObj) {
  const txtShadow = getTextByPathList(node, ['a:rPr', 'a:effectLst', 'a:outerShdw'])
  if (txtShadow) {
    const shadow = getShadow(txtShadow, warpObj)
    if (shadow) {
      const { h, v, blur, color } = shadow
      if (!isNaN(v) && !isNaN(h)) {
        return h + 'pt ' + v + 'pt ' + (blur ? blur + 'pt' : '') + ' ' + color
      }
    }
  }
  return ''
}
