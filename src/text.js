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

    const listInfo = getListInfo(pNode)
    if (listInfo) {
      const nextKey = getListKey(listInfo)
      if (!currentListState || currentListState.key !== nextKey) {
        if (currentListState) text += `</${currentListState.tag}>`
        text += `<${listInfo.tag} style="list-style: none; padding-left: 0; margin: 0;">`
        currentListState = {
          key: nextKey,
          tag: listInfo.tag,
          listInfo,
          counter: listInfo.kind === 'autoNum' ? listInfo.startAt : null,
        }
      }

      const marker = getListMarker(currentListState)
      const bulletStyle = getListMarkerStyle(currentListState.listInfo)
      const indent = (listInfo.lvl - 1) * 1.5
      text += `<li style="text-align: ${align}; margin-left: ${indent}em;"><span style="${bulletStyle}">${marker}</span>`
    }
    else {
      if (currentListState) {
        text += `</${currentListState.tag}>`
        currentListState = null
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

export function getListInfo(node) {
  const pPrNode = node['a:pPr']
  if (!pPrNode) return null
  if (pPrNode['a:buNone']) return null

  let lvl = 1
  const lvlNode = getTextByPathList(pPrNode, ['attrs', 'lvl'])
  if (lvlNode !== undefined) lvl = parseInt(lvlNode) + 1

  if (pPrNode['a:buChar']) {
    const char = getTextByPathList(pPrNode, ['a:buChar', 'attrs', 'char']) || '•'
    const font = getTextByPathList(pPrNode, ['a:buFont', 'attrs', 'typeface']) || ''
    return {
      kind: 'char',
      tag: 'ul',
      lvl,
      char,
      font,
    }
  }

  if (pPrNode['a:buAutoNum']) {
    const autoNumNode = pPrNode['a:buAutoNum']
    const numType = getTextByPathList(autoNumNode, ['attrs', 'type']) || 'arabicPeriod'
    const startAtRaw = getTextByPathList(autoNumNode, ['attrs', 'startAt'])
    const startAt = startAtRaw ? parseInt(startAtRaw) : 1
    const font = getTextByPathList(pPrNode, ['a:buFont', 'attrs', 'typeface']) || ''
    return {
      kind: 'autoNum',
      tag: 'ol',
      lvl,
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
  const suffix = numType.includes('ParenR') ? ')' : (numType.includes('Period') ? '.' : '')
  const bothParen = numType.includes('ParenBoth')

  let core
  if (numType.includes('alphaLc')) core = toAlpha(n, false)
  else if (numType.includes('alphaUc')) core = toAlpha(n, true)
  else if (numType.includes('romanLc')) core = toRoman(n, false)
  else if (numType.includes('romanUc')) core = toRoman(n, true)
  else core = String(n)

  if (bothParen) return `(${core})`
  return `${core}${suffix || '.'}`
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
