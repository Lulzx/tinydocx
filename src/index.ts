import { deflateRawSync } from 'node:zlib'

export interface TextOptions {
  align?: 'left' | 'center' | 'right'
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strikethrough?: boolean
  code?: boolean
  color?: string
  font?: string
  size?: number
}

export interface TableOptions { colWidths?: number[] }
export interface ImageOptions { width: number; height: number }

export interface TextRun {
  text: string
  opts?: TextOptions
  link?: string
}

export interface DocContext {
  heading(str: string, level: 1 | 2 | 3 | 4 | 5 | 6): void
  paragraph(str: string, opts?: TextOptions): void
  text(str: string, size: number, opts?: TextOptions): void
  lineBreak(): void
  horizontalRule(): void
  list(items: string[], ordered?: boolean): void
  table(rows: string[][], opts?: TableOptions): void
  link(text: string, url: string, opts?: TextOptions): void
  image(data: Uint8Array, opts: ImageOptions): void
  pageNumber(): void
}

export interface DOCXBuilder {
  content(fn: (ctx: DocContext) => void): DOCXBuilder
  header(fn: (ctx: DocContext) => void): DOCXBuilder
  footer(fn: (ctx: DocContext) => void): DOCXBuilder
  build(): Uint8Array
}

export interface ODTBuilder {
  content(fn: (ctx: DocContext) => void): ODTBuilder
  header(fn: (ctx: DocContext) => void): ODTBuilder
  footer(fn: (ctx: DocContext) => void): ODTBuilder
  build(): Uint8Array
}

interface ListItem {
  runs: TextRun[]
  children?: { items: ListItem[]; ordered: boolean }
}

type DocElement =
  | { type: 'heading'; text: string; level: 1 | 2 | 3 | 4 | 5 | 6 }
  | { type: 'paragraph'; text: string; opts?: TextOptions }
  | { type: 'richParagraph'; runs: TextRun[]; align?: 'left' | 'center' | 'right' }
  | { type: 'text'; text: string; size: number; opts?: TextOptions }
  | { type: 'lineBreak' }
  | { type: 'horizontalRule' }
  | { type: 'list'; items: string[]; ordered: boolean }
  | { type: 'richList'; items: ListItem[]; ordered: boolean }
  | { type: 'table'; rows: string[][]; opts?: TableOptions }
  | { type: 'richTable'; rows: TextRun[][][]; opts?: TableOptions }
  | { type: 'link'; text: string; url: string; opts?: TextOptions; rId: string }
  | { type: 'image'; data: Uint8Array; opts: ImageOptions; rId: string; docPrId: number }
  | { type: 'blockquote'; elements: DocElement[] }
  | { type: 'codeBlock'; code: string; language?: string }
  | { type: 'pageNumber' }

type InlineToken =
  | { tag: 'text'; content: string }
  | { tag: 'bold'; children: InlineToken[] }
  | { tag: 'italic'; children: InlineToken[] }
  | { tag: 'code'; content: string }
  | { tag: 'strike'; children: InlineToken[] }
  | { tag: 'link'; children: InlineToken[]; href: string }
  | { tag: 'image'; alt: string; src: string }

type BlockToken =
  | { tag: 'h'; level: 1 | 2 | 3 | 4 | 5 | 6; content: InlineToken[] }
  | { tag: 'p'; content: InlineToken[] }
  | { tag: 'pre'; code: string; lang?: string }
  | { tag: 'quote'; blocks: BlockToken[] }
  | { tag: 'hr' }
  | { tag: 'ul'; items: ListNodeItem[] }
  | { tag: 'ol'; items: ListNodeItem[] }
  | { tag: 'table'; head: InlineToken[][]; body: InlineToken[][][] }

type ListNodeItem = {
  content: InlineToken[]
  nested?: { ordered: boolean; items: ListNodeItem[] }
}

type Result<T> = { ok: true; val: T; rest: string } | { ok: false }
type Parser<T> = (s: string) => Result<T>

const ok = <T>(val: T, rest: string): Result<T> => ({ ok: true, val, rest })
const fail = <T>(): Result<T> => ({ ok: false })

const lit = (match: string): Parser<string> => s =>
  s.startsWith(match) ? ok(match, s.slice(match.length)) : fail()

const re = (pattern: RegExp): Parser<string> => s => {
  const m = s.match(new RegExp('^' + pattern.source))
  return m ? ok(m[0], s.slice(m[0].length)) : fail()
}

const map = <A, B>(p: Parser<A>, f: (a: A) => B): Parser<B> => s => {
  const r = p(s)
  return r.ok ? ok(f(r.val), r.rest) : fail()
}

const seq = <A, B>(p1: Parser<A>, p2: Parser<B>): Parser<[A, B]> => s => {
  const r1 = p1(s)
  if (!r1.ok) return fail()
  const r2 = p2(r1.rest)
  return r2.ok ? ok([r1.val, r2.val], r2.rest) : fail()
}

const alt = <T>(...ps: Parser<T>[]): Parser<T> => s =>
  ps.reduce<Result<T>>((acc, p) => acc.ok ? acc : p(s), fail())

const MAX_PARSE_DEPTH = 1000

const many = <T>(p: Parser<T>): Parser<T[]> => {
  const go = (s: string, acc: T[], depth: number): Result<T[]> => {
    if (depth >= MAX_PARSE_DEPTH) return ok(acc, s)
    const r = p(s)
    return r.ok ? go(r.rest, [...acc, r.val], depth + 1) : ok(acc, s)
  }
  return s => go(s, [], 0)
}

const lazy = <T>(f: () => Parser<T>): Parser<T> => s => f()(s)

const between = <T>(open: string, close: string, p: Parser<T>): Parser<T> => s => {
  const r1 = lit(open)(s)
  if (!r1.ok) return fail()
  const r2 = p(r1.rest)
  if (!r2.ok) return fail()
  const r3 = lit(close)(r2.rest)
  return r3.ok ? ok(r2.val, r3.rest) : fail()
}

const takeWhile = (pred: (c: string) => boolean): Parser<string> => {
  const go = (s: string, i: number): Result<string> =>
    i < s.length && pred(s[i]) ? go(s, i + 1) : i > 0 ? ok(s.slice(0, i), s.slice(i)) : fail()
  return s => go(s, 0)
}

const anyChar: Parser<string> = s => s.length > 0 ? ok(s[0], s.slice(1)) : fail()

const inlineSpecials = new Set(['*', '_', '`', '~', '[', '!', '\\'])

const escapeP: Parser<InlineToken> = map(seq(lit('\\'), anyChar), ([, c]) => ({ tag: 'text', content: c }))

const codeP: Parser<InlineToken> = map(between('`', '`', takeWhile(c => c !== '`')), content => ({ tag: 'code', content }))

const inlineTokens: Parser<InlineToken[]> = lazy(() => many(inlineElem))

const untilInline = (delim: string): Parser<InlineToken[]> => {
  const go = (s: string, acc: InlineToken[]): Result<InlineToken[]> => {
    if (s.length === 0 || s.startsWith(delim)) return ok(acc, s)
    const r = inlineElem(s)
    return r.ok ? go(r.rest, [...acc, r.val]) : ok(acc, s)
  }
  return s => go(s, [])
}

const boldP: Parser<InlineToken> = alt(
  map(between('**', '**', lazy(() => untilInline('**'))), children => ({ tag: 'bold', children })),
  map(between('__', '__', lazy(() => untilInline('__'))), children => ({ tag: 'bold', children }))
)

const strikeP: Parser<InlineToken> = map(
  between('~~', '~~', lazy(() => untilInline('~~'))),
  children => ({ tag: 'strike', children })
)

const italicP: Parser<InlineToken> = s => {
  if (s.length < 3) return fail()
  if (s[0] !== '*' && s[0] !== '_') return fail()
  if (s[1] === s[0]) return fail()
  const marker = s[0]
  const end = s.indexOf(marker, 1)
  if (end <= 1) return fail()
  const children = parseInlineTokens(s.slice(1, end))
  return ok({ tag: 'italic', children }, s.slice(end + 1))
}

const imageP: Parser<InlineToken> = s => {
  if (!s.startsWith('![')) return fail()
  const altEnd = s.indexOf(']', 2)
  if (altEnd === -1 || s[altEnd + 1] !== '(') return fail()
  const srcEnd = s.indexOf(')', altEnd + 2)
  if (srcEnd === -1) return fail()
  return ok({ tag: 'image', alt: s.slice(2, altEnd), src: s.slice(altEnd + 2, srcEnd) }, s.slice(srcEnd + 1))
}

const linkP: Parser<InlineToken> = s => {
  if (s[0] !== '[') return fail()
  const textEnd = s.indexOf(']', 1)
  if (textEnd === -1 || s[textEnd + 1] !== '(') return fail()
  const hrefEnd = s.indexOf(')', textEnd + 2)
  if (hrefEnd === -1) return fail()
  const children = parseInlineTokens(s.slice(1, textEnd))
  return ok({ tag: 'link', children, href: s.slice(textEnd + 2, hrefEnd) }, s.slice(hrefEnd + 1))
}

const plainTextP: Parser<InlineToken> = map(takeWhile(c => !inlineSpecials.has(c)), content => ({ tag: 'text', content }))

const singleCharP: Parser<InlineToken> = map(anyChar, content => ({ tag: 'text', content }))

const inlineElem: Parser<InlineToken> = alt(escapeP, codeP, boldP, strikeP, italicP, imageP, linkP, plainTextP, singleCharP)

const parseInlineTokens = (s: string): InlineToken[] => {
  const r = inlineTokens(s)
  return r.ok ? r.val : [{ tag: 'text', content: s }]
}

const mergeTextTokens = (tokens: InlineToken[]): InlineToken[] => {
  const result: InlineToken[] = []
  for (const t of tokens) {
    const last = result[result.length - 1]
    if (t.tag === 'text' && last?.tag === 'text') {
      result[result.length - 1] = { tag: 'text', content: last.content + t.content }
    } else {
      result.push(t)
    }
  }
  return result
}

const tokensToRuns = (tokens: InlineToken[], inherited: TextOptions = {}): TextRun[] =>
  mergeTextTokens(tokens).flatMap((t): TextRun[] => {
    const hasOpts = Object.keys(inherited).length > 0
    switch (t.tag) {
      case 'text': return [{ text: t.content, opts: hasOpts ? { ...inherited } : undefined }]
      case 'bold': return tokensToRuns(t.children, { ...inherited, bold: true })
      case 'italic': return tokensToRuns(t.children, { ...inherited, italic: true })
      case 'strike': return tokensToRuns(t.children, { ...inherited, strikethrough: true })
      case 'code': return [{ text: t.content, opts: { ...inherited, code: true } }]
      case 'link': return tokensToRuns(t.children, { ...inherited, color: '#0563C1', underline: true }).map(r => ({ ...r, link: t.href }))
      case 'image': return []
    }
  })

const tokensToPlainText = (tokens: InlineToken[]): string =>
  tokens.map(t => {
    switch (t.tag) {
      case 'text': case 'code': return t.content
      case 'bold': case 'italic': case 'strike': case 'link': return tokensToPlainText(t.children)
      case 'image': return t.alt
    }
  }).join('')

type LineState = { lines: readonly string[]; idx: number }

const peek = (s: LineState): string | undefined => s.lines[s.idx]
const advance = (s: LineState): LineState => ({ ...s, idx: s.idx + 1 })
const isDone = (s: LineState): boolean => s.idx >= s.lines.length

const parseCodeBlock = (state: LineState): [BlockToken, LineState] | null => {
  const line = peek(state)
  if (!line?.startsWith('```')) return null
  const lang = line.slice(3).trim() || undefined
  const collectCode = (s: LineState, acc: string[]): [string[], LineState] => {
    const l = peek(s)
    if (l === undefined || l.startsWith('```')) return [acc, advance(s)]
    return collectCode(advance(s), [...acc, l])
  }
  const [codeLines, next] = collectCode(advance(state), [])
  return [{ tag: 'pre', code: codeLines.join('\n'), lang }, next]
}

const parseHeading = (state: LineState): [BlockToken, LineState] | null => {
  const line = peek(state)
  if (!line || !/^#{1,6}\s/.test(line)) return null
  const m = line.match(/^(#{1,6})\s+(.*)$/)
  if (!m) return null
  const level = Math.min(m[1].length, 6) as 1 | 2 | 3 | 4 | 5 | 6
  return [{ tag: 'h', level, content: parseInlineTokens(m[2]) }, advance(state)]
}

const parseHr = (state: LineState): [BlockToken, LineState] | null => {
  const line = peek(state)
  if (!line || !/^(-{3,}|\*{3,}|_{3,})$/.test(line.trim())) return null
  return [{ tag: 'hr' }, advance(state)]
}

const parseQuote = (state: LineState): [BlockToken, LineState] | null => {
  const line = peek(state)
  if (!line?.startsWith('>')) return null
  const stripQuote = (l: string): string => l.startsWith('> ') ? l.slice(2) : l.slice(1)
  const isQuoteLine = (l: string | undefined): boolean => l?.startsWith('>') ?? false
  const collectQuote = (s: LineState, acc: string[]): [string[], LineState] => {
    const l = peek(s)
    if (!isQuoteLine(l)) return [acc, s]
    return collectQuote(advance(s), [...acc, stripQuote(l!)])
  }
  const [quoteLines, next] = collectQuote(state, [])
  return [{ tag: 'quote', blocks: parseBlocks({ lines: quoteLines, idx: 0 }) }, next]
}

const parseTableRow = (r: string): InlineToken[][] =>
  r.slice(1, -1).split('|').map(c => parseInlineTokens(c.trim()))

const parseTable = (state: LineState): [BlockToken, LineState] | null => {
  const line = peek(state)
  if (!line || !/^\|.+\|$/.test(line)) return null
  const head = parseTableRow(line)
  let next = advance(state)
  const sep = peek(next)
  if (sep && /^\|[-:| ]+\|$/.test(sep)) next = advance(next)
  const collectBody = (s: LineState, acc: InlineToken[][][]): [InlineToken[][][], LineState] => {
    const l = peek(s)
    if (!l || !/^\|.+\|$/.test(l)) return [acc, s]
    return collectBody(advance(s), [...acc, parseTableRow(l)])
  }
  const [body, final] = collectBody(next, [])
  return [{ tag: 'table', head, body }, final]
}

const parseList = (state: LineState): [BlockToken, LineState] | null => {
  const line = peek(state)
  if (!line) return null
  const isUnordered = /^(\s*)[-*]\s/.test(line)
  const isOrdered = /^(\s*)\d+\.\s/.test(line)
  if (!isUnordered && !isOrdered) return null

  const getIndent = (l: string): number => l.match(/^(\s*)/)?.[1].length ?? 0
  const getMarker = (l: string): 'ul' | 'ol' | null =>
    /^\s*[-*]\s/.test(l) ? 'ul' : /^\s*\d+\.\s/.test(l) ? 'ol' : null
  const stripMarker = (l: string): string => l.replace(/^\s*[-*]\s|^\s*\d+\.\s/, '')

  const parseItems = (s: LineState, indent: number): [ListNodeItem[], LineState] => {
    const collectItems = (st: LineState, acc: ListNodeItem[]): [ListNodeItem[], LineState] => {
      const l = peek(st)
      if (!l) return [acc, st]
      const lIndent = getIndent(l)
      const marker = getMarker(l)
      if (lIndent < indent) return [acc, st]
      if (lIndent === indent && marker) {
        const content = parseInlineTokens(stripMarker(l))
        let next = advance(st)
        let nested: ListNodeItem['nested'] = undefined
        const nextLine = peek(next)
        if (nextLine && getIndent(nextLine) > indent && getMarker(nextLine)) {
          const [nestedItems, afterNested] = parseItems(next, getIndent(nextLine))
          nested = { ordered: getMarker(nextLine) === 'ol', items: nestedItems }
          next = afterNested
        }
        return collectItems(next, [...acc, { content, nested }])
      }
      if (lIndent > indent) return collectItems(advance(st), acc)
      return [acc, st]
    }
    return collectItems(s, [])
  }

  const indent = getIndent(line)
  const [items, next] = parseItems(state, indent)
  return [isOrdered ? { tag: 'ol', items } : { tag: 'ul', items }, next]
}

const isSpecialLine = (line: string): boolean =>
  /^#{1,6}\s/.test(line) || /^(-{3,}|\*{3,}|_{3,})$/.test(line.trim()) ||
  line.startsWith('>') || line.startsWith('```') || /^\|.+\|$/.test(line) ||
  /^(\s*)[-*]\s/.test(line) || /^(\s*)\d+\.\s/.test(line)

const parseParagraph = (state: LineState): [BlockToken, LineState] | null => {
  const line = peek(state)
  if (!line || line.trim() === '') return null
  const lines: string[] = [line]
  let cur = advance(state)
  while (!isDone(cur)) {
    const l = peek(cur)
    if (!l || l.trim() === '' || isSpecialLine(l)) break
    lines.push(l)
    cur = advance(cur)
  }
  return [{ tag: 'p', content: parseInlineTokens(lines.join(' ')) }, cur]
}

const parseBlocks = (state: LineState): BlockToken[] => {
  const go = (s: LineState, acc: BlockToken[]): BlockToken[] => {
    if (isDone(s)) return acc
    const line = peek(s)
    if (line?.trim() === '') return go(advance(s), acc)
    const parsers = [parseCodeBlock, parseHeading, parseHr, parseQuote, parseTable, parseList, parseParagraph]
    for (const parser of parsers) {
      const result = parser(s)
      if (result) {
        const [block, next] = result
        return go(next, [...acc, block])
      }
    }
    return go(advance(s), acc)
  }
  return go(state, [])
}

const listNodeToListItem = (node: ListNodeItem): ListItem => ({
  runs: tokensToRuns(node.content),
  children: node.nested ? { ordered: node.nested.ordered, items: node.nested.items.map(listNodeToListItem) } : undefined
})

const blockToElement = (b: BlockToken): DocElement[] => {
  switch (b.tag) {
    case 'h': return [{ type: 'heading', text: tokensToPlainText(b.content), level: b.level }]
    case 'p': return [{ type: 'richParagraph', runs: tokensToRuns(b.content) }]
    case 'pre': return [{ type: 'codeBlock', code: b.code, language: b.lang }]
    case 'hr': return [{ type: 'horizontalRule' }]
    case 'quote': return [{ type: 'blockquote', elements: b.blocks.flatMap(blockToElement) }]
    case 'ul': return [{ type: 'richList', ordered: false, items: b.items.map(listNodeToListItem) }]
    case 'ol': return [{ type: 'richList', ordered: true, items: b.items.map(listNodeToListItem) }]
    case 'table': return [{ type: 'richTable', rows: [b.head.map(h => tokensToRuns(h)), ...b.body.map(r => r.map(c => tokensToRuns(c)))] }]
  }
}

const parseMarkdownAST = (md: string): BlockToken[] =>
  parseBlocks({ lines: md.replace(/\r\n/g, '\n').split('\n'), idx: 0 })

const astToElements = (blocks: BlockToken[]): DocElement[] => blocks.flatMap(blockToElement)

const crc32Table = Array.from({ length: 256 }, (_, i) =>
  Array.from({ length: 8 }).reduce<number>((c) => (c >>> 1) ^ (c & 1 ? 0xedb88320 : 0), i)
)

const crc32 = (data: Uint8Array): number =>
  (data.reduce((crc, byte) => crc32Table[(crc ^ byte) & 0xff] ^ (crc >>> 8), 0xffffffff) ^ 0xffffffff) >>> 0

const escapeXml = (str: string): string =>
  str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')

const COURIER = 'Courier New'
const OOXML = 'http://schemas.openxmlformats.org/'
const OD = 'urn:oasis:names:tc:opendocument:xmlns:'
const NS = {
  w: OOXML + 'wordprocessingml/2006/main', r: OOXML + 'officeDocument/2006/relationships',
  rel: OOXML + 'package/2006/relationships', ct: OOXML + 'package/2006/content-types',
  wp: OOXML + 'drawingml/2006/wordprocessingDrawing', a: OOXML + 'drawingml/2006/main',
  pic: OOXML + 'drawingml/2006/picture', manifest: OD + 'manifest:1.0',
  office: OD + 'office:1.0', text: OD + 'text:1.0', style: OD + 'style:1.0',
  fo: OD + 'xsl-fo-compatible:1.0', table: OD + 'table:1.0',
  draw: OD + 'drawing:1.0', svg: OD + 'svg-compatible:1.0',
  xlink: 'http://www.w3.org/1999/xlink',
}

type ZipEntry = { name: string; data: Uint8Array }

const createZip = (files: ZipEntry[]): Uint8Array => {
  const enc = new TextEncoder()
  const prepared = files.map(f => {
    const name = enc.encode(f.name)
    const crcVal = crc32(f.data)
    const compressed = deflateRawSync(f.data)
    const useDeflate = compressed.length < f.data.length
    return { name, data: f.data, compressed: useDeflate ? compressed : f.data, crc: crcVal, method: useDeflate ? 8 : 0 }
  })

  let totalSize = 0
  const offsets: number[] = []
  for (const p of prepared) {
    offsets.push(totalSize)
    totalSize += 30 + p.name.length + p.compressed.length
  }
  const centralStart = totalSize
  for (const p of prepared) totalSize += 46 + p.name.length
  totalSize += 22

  const buf = new Uint8Array(totalSize)
  const view = new DataView(buf.buffer)
  let pos = 0

  for (let i = 0; i < prepared.length; i++) {
    const p = prepared[i]
    view.setUint32(pos, 0x04034b50, true); view.setUint16(pos + 4, 20, true)
    view.setUint16(pos + 8, p.method, true)
    view.setUint32(pos + 14, p.crc, true)
    view.setUint32(pos + 18, p.compressed.length, true)
    view.setUint32(pos + 22, p.data.length, true)
    view.setUint16(pos + 26, p.name.length, true)
    buf.set(p.name, pos + 30)
    buf.set(p.compressed, pos + 30 + p.name.length)
    pos += 30 + p.name.length + p.compressed.length
  }

  for (let i = 0; i < prepared.length; i++) {
    const p = prepared[i]
    view.setUint32(pos, 0x02014b50, true); view.setUint16(pos + 4, 20, true); view.setUint16(pos + 6, 20, true)
    view.setUint16(pos + 10, p.method, true)
    view.setUint32(pos + 16, p.crc, true)
    view.setUint32(pos + 20, p.compressed.length, true)
    view.setUint32(pos + 24, p.data.length, true)
    view.setUint16(pos + 28, p.name.length, true)
    view.setUint32(pos + 42, offsets[i], true)
    buf.set(p.name, pos + 46)
    pos += 46 + p.name.length
  }

  const centralSize = pos - centralStart
  view.setUint32(pos, 0x06054b50, true)
  view.setUint16(pos + 8, prepared.length, true)
  view.setUint16(pos + 10, prepared.length, true)
  view.setUint32(pos + 12, centralSize, true)
  view.setUint32(pos + 16, centralStart, true)

  return buf
}

interface BuildContext {
  elements: DocElement[]
  hyperlinks: { url: string; rId: string }[]
  images: { data: Uint8Array; rId: string; ext: string }[]
  nextRId: number
  nextDocPrId: { value: number }
}

const detectImageType = (data: Uint8Array): string => {
  if (data.length < 4) return 'png'
  if (data[0] === 0x89 && data[1] === 0x50 && data[2] === 0x4e && data[3] === 0x47) return 'png'
  if (data[0] === 0xff && data[1] === 0xd8) return 'jpeg'
  if (data[0] === 0x47 && data[1] === 0x49 && data[2] === 0x46) return 'gif'
  if (data.length >= 12 && data[0] === 0x52 && data[1] === 0x49 && data[2] === 0x46 && data[3] === 0x46 &&
      data[8] === 0x57 && data[9] === 0x45 && data[10] === 0x42 && data[11] === 0x50) return 'webp'
  return 'png'
}

const createContext = (ctx: BuildContext): DocContext => ({
  heading: (str, level) => { ctx.elements.push({ type: 'heading', text: str, level }) },
  paragraph: (str, opts) => { ctx.elements.push({ type: 'paragraph', text: str, opts }) },
  text: (str, size, opts) => { ctx.elements.push({ type: 'text', text: str, size, opts }) },
  lineBreak: () => { ctx.elements.push({ type: 'lineBreak' }) },
  horizontalRule: () => { ctx.elements.push({ type: 'horizontalRule' }) },
  list: (items, ordered = false) => { ctx.elements.push({ type: 'list', items, ordered }) },
  table: (rows, opts) => { ctx.elements.push({ type: 'table', rows, opts }) },
  link: (text, url, opts) => {
    const rId = `rId${ctx.nextRId++}`
    ctx.hyperlinks.push({ url, rId })
    ctx.elements.push({ type: 'link', text, url, opts, rId })
  },
  image: (data, opts) => {
    const rId = `rId${ctx.nextRId++}`
    const docPrId = ctx.nextDocPrId.value++
    const ext = detectImageType(data)
    ctx.images.push({ data, rId, ext })
    ctx.elements.push({ type: 'image', data, opts, rId, docPrId })
  },
  pageNumber: () => { ctx.elements.push({ type: 'pageNumber' }) }
})

const buildDocxRunProps = (opts?: TextOptions, size?: number): string => {
  const parts: string[] = []
  const font = opts?.code ? COURIER : opts?.font
  if (font) parts.push(`<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/>`)
  const sz = size || opts?.size
  if (sz) parts.push(`<w:sz w:val="${sz * 2}"/><w:szCs w:val="${sz * 2}"/>`)
  if (opts?.bold) parts.push('<w:b/>')
  if (opts?.italic) parts.push('<w:i/>')
  if (opts?.underline) parts.push('<w:u w:val="single"/>')
  if (opts?.strikethrough) parts.push('<w:strike/>')
  if (opts?.code) parts.push('<w:shd w:val="clear" w:color="auto" w:fill="E8E8E8"/>')
  if (opts?.color) parts.push(`<w:color w:val="${opts.color.replace('#', '')}"/>`)
  return parts.length > 0 ? `<w:rPr>${parts.join('')}</w:rPr>` : ''
}

const getAlignment = (align?: string): string =>
  align === 'center' ? 'center' : align === 'right' ? 'right' : align === 'justify' ? 'both' : 'left'

const parseInline = (text: string): TextRun[] => tokensToRuns(parseInlineTokens(text))

const buildDocxRuns = (runs: TextRun[]): string =>
  runs.map(run => `<w:r>${buildDocxRunProps(run.opts)}<w:t>${escapeXml(run.text)}</w:t></w:r>`).join('')

const buildDocxListItems = (items: ListItem[], ordered: boolean, level: number): string =>
  items.flatMap(item => [
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${ordered ? 2 : 1}"/></w:numPr></w:pPr>${buildDocxRuns(item.runs)}</w:p>`,
    ...(item.children ? [buildDocxListItems(item.children.items, item.children.ordered, level + 1)] : [])
  ]).join('')

const HEADING_SIZES: Record<1 | 2 | 3 | 4 | 5 | 6, number> = { 1: 48, 2: 36, 3: 28, 4: 24, 5: 20, 6: 18 }

const TABLE_BORDERS = '<w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/>'

const elementToDocx = (el: DocElement): string => {
  switch (el.type) {
    case 'heading': {
      const sz = HEADING_SIZES[el.level]
      return `<w:p><w:pPr><w:pStyle w:val="Heading${el.level}"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr><w:t>${escapeXml(el.text)}</w:t></w:r></w:p>`
    }
    case 'paragraph':
      return `<w:p><w:pPr><w:jc w:val="${getAlignment(el.opts?.align)}"/></w:pPr><w:r>${buildDocxRunProps(el.opts)}<w:t>${escapeXml(el.text)}</w:t></w:r></w:p>`
    case 'text':
      return `<w:p><w:pPr><w:jc w:val="${getAlignment(el.opts?.align)}"/></w:pPr><w:r>${buildDocxRunProps(el.opts, el.size)}<w:t>${escapeXml(el.text)}</w:t></w:r></w:p>`
    case 'lineBreak': return '<w:p/>'
    case 'horizontalRule':
      return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>'
    case 'list':
      return el.items.map(item => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${el.ordered ? 2 : 1}"/></w:numPr></w:pPr><w:r><w:t>${escapeXml(item)}</w:t></w:r></w:p>`).join('')
    case 'table': {
      const grid = el.opts?.colWidths ? `<w:tblGrid>${el.opts.colWidths.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>` : ''
      const rows = el.rows.map(row => `<w:tr>${row.map((cell, i) => `<w:tc>${el.opts?.colWidths?.[i] ? `<w:tcPr><w:tcW w:w="${el.opts.colWidths[i]}" w:type="dxa"/></w:tcPr>` : ''}<w:p><w:r><w:t>${escapeXml(cell)}</w:t></w:r></w:p></w:tc>`).join('')}</w:tr>`).join('')
      return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${TABLE_BORDERS}</w:tblBorders></w:tblPr>${grid}${rows}</w:tbl>`
    }
    case 'link':
      return `<w:p><w:hyperlink r:id="${el.rId}"><w:r>${buildDocxRunProps({ ...el.opts, color: el.opts?.color || '0563C1', underline: true })}<w:t>${escapeXml(el.text)}</w:t></w:r></w:hyperlink></w:p>`
    case 'image': {
      const cx = Math.round(el.opts.width * 914400), cy = Math.round(el.opts.height * 914400)
      return `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="${NS.wp}"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${el.docPrId}" name="Image${el.docPrId}"/><a:graphic xmlns:a="${NS.a}"><a:graphicData uri="${NS.pic}"><pic:pic xmlns:pic="${NS.pic}"><pic:nvPicPr><pic:cNvPr id="${el.docPrId}" name="Image${el.docPrId}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${el.rId}" xmlns:r="${NS.r}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
    }
    case 'pageNumber':
      return '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    case 'richParagraph':
      return `<w:p><w:pPr><w:jc w:val="${getAlignment(el.align)}"/></w:pPr>${buildDocxRuns(el.runs)}</w:p>`
    case 'richList': return buildDocxListItems(el.items, el.ordered, 0)
    case 'richTable': {
      const grid = el.opts?.colWidths ? `<w:tblGrid>${el.opts.colWidths.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>` : ''
      const rows = el.rows.map(row => `<w:tr>${row.map((cell, i) => `<w:tc>${el.opts?.colWidths?.[i] ? `<w:tcPr><w:tcW w:w="${el.opts.colWidths[i]}" w:type="dxa"/></w:tcPr>` : ''}<w:p>${buildDocxRuns(cell)}</w:p></w:tc>`).join('')}</w:tr>`).join('')
      return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${TABLE_BORDERS}</w:tblBorders></w:tblPr>${grid}${rows}</w:tbl>`
    }
    case 'blockquote': {
      const quoteStyle = '<w:pPr><w:ind w:left="720"/><w:pBdr><w:left w:val="single" w:sz="18" w:space="4" w:color="CCCCCC"/></w:pBdr></w:pPr>'
      return el.elements.map(child => {
        if (child.type === 'richParagraph') return `<w:p>${quoteStyle}${buildDocxRuns(child.runs)}</w:p>`
        return elementToDocx(child)
      }).join('')
    }
    case 'codeBlock':
      return el.code.split('\n').map(line => `<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${COURIER}" w:hAnsi="${COURIER}"/></w:rPr><w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r></w:p>`).join('')
  }
}

const buildDocxBody = (elements: DocElement[]): string => elements.map(elementToDocx).join('')

const generateDocxDocument = (elements: DocElement[], headerRId?: string, footerRId?: string): string => {
  const body = buildDocxBody(elements)
  const sectPr = `<w:sectPr>${headerRId ? `<w:headerReference w:type="default" r:id="${headerRId}"/>` : ''}${footerRId ? `<w:footerReference w:type="default" r:id="${footerRId}"/>` : ''}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>`
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}">\n<w:body>\n${body}\n${sectPr}\n</w:body>\n</w:document>`
}

const generateDocxHeader = (elements: DocElement[], hasLinksOrImages: boolean): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr xmlns:w="${NS.w}"${hasLinksOrImages ? ` xmlns:r="${NS.r}"` : ''}>\n${buildDocxBody(elements)}\n</w:hdr>`

const generateDocxFooter = (elements: DocElement[], hasLinksOrImages: boolean): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:ftr xmlns:w="${NS.w}"${hasLinksOrImages ? ` xmlns:r="${NS.r}"` : ''}>\n${buildDocxBody(elements)}\n</w:ftr>`

const generateDocxStyles = (): string => {
  const h = [1, 2, 3, 4, 5, 6].map(i => `<w:style w:type="paragraph" w:styleId="Heading${i}"><w:name w:val="Heading ${i}"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="${HEADING_SIZES[i as 1|2|3|4|5|6]}"/></w:rPr></w:style>`).join('')
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${NS.w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${h}</w:styles>`
}

const BULLET_CHARS = ['•', '◦', '▪', '•', '◦']
const generateDocxNumbering = (): string => {
  const bulletLevels = Array.from({ length: 5 }, (_, i) =>
    `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${BULLET_CHARS[i]}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`
  ).join('')
  const decimalLevels = Array.from({ length: 5 }, (_, i) =>
    `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%${i + 1}."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`
  ).join('')
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${NS.w}"><w:abstractNum w:abstractNumId="0">${bulletLevels}</w:abstractNum><w:abstractNum w:abstractNumId="1">${decimalLevels}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
}

const generateDocxContentTypes = (hasLists: boolean, hasHeader: boolean, hasFooter: boolean, imageExts: string[]): string => {
  const imgDefaults = Array.from(new Set(imageExts)).map(ext => `<Default Extension="${ext}" ContentType="${ext === 'jpeg' ? 'image/jpeg' : ext === 'gif' ? 'image/gif' : 'image/png'}"/>`).join('')
  const overrides = `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${hasLists ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : ''}${hasHeader ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : ''}${hasFooter ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : ''}`
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${NS.ct}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${imgDefaults}${overrides}</Types>`
}

const generateDocxRels = (hasLists: boolean, hasHeader: boolean, hasFooter: boolean, hyperlinks: { url: string; rId: string }[], images: { rId: string; ext: string }[], imageOffset: number): string => {
  const rels = [`<Relationship Id="rId1" Type="${NS.r}/styles" Target="styles.xml"/>`,
    ...(hasLists ? [`<Relationship Id="rId2" Type="${NS.r}/numbering" Target="numbering.xml"/>`] : []),
    ...(hasHeader ? [`<Relationship Id="rIdHeader" Type="${NS.r}/header" Target="header1.xml"/>`] : []),
    ...(hasFooter ? [`<Relationship Id="rIdFooter" Type="${NS.r}/footer" Target="footer1.xml"/>`] : []),
    ...hyperlinks.map(l => `<Relationship Id="${l.rId}" Type="${NS.r}/hyperlink" Target="${escapeXml(l.url)}" TargetMode="External"/>`),
    ...images.map((img, i) => `<Relationship Id="${img.rId}" Type="${NS.r}/image" Target="media/image${imageOffset + i + 1}.${img.ext}"/>`)
  ].join('')
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}">${rels}</Relationships>`
}

const generatePartRels = (hyperlinks: { url: string; rId: string }[], images: { rId: string; ext: string }[], imageOffset: number): string => {
  const rels = [
    ...hyperlinks.map(l => `<Relationship Id="${l.rId}" Type="${NS.r}/hyperlink" Target="${escapeXml(l.url)}" TargetMode="External"/>`),
    ...images.map((img, i) => `<Relationship Id="${img.rId}" Type="${NS.r}/image" Target="media/image${imageOffset + i + 1}.${img.ext}"/>`)
  ].join('')
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}">${rels}</Relationships>`
}

const DOCX_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${NS.r}/officeDocument" Target="word/document.xml"/></Relationships>`

/**
 * Create a new DOCX document
 */
export function docx(): DOCXBuilder {
  const docPrIdCounter = { value: 1 }
  const mainCtx: BuildContext = { elements: [], hyperlinks: [], images: [], nextRId: 10, nextDocPrId: docPrIdCounter }
  const headerCtx: BuildContext = { elements: [], hyperlinks: [], images: [], nextRId: 100, nextDocPrId: docPrIdCounter }
  const footerCtx: BuildContext = { elements: [], hyperlinks: [], images: [], nextRId: 200, nextDocPrId: docPrIdCounter }
  let hasHeader = false, hasFooter = false

  const builder: DOCXBuilder = {
    content(fn) { fn(createContext(mainCtx)); return builder },
    header(fn) { hasHeader = true; fn(createContext(headerCtx)); return builder },
    footer(fn) { hasFooter = true; fn(createContext(footerCtx)); return builder },
    build() {
      const enc = new TextEncoder()
      const hasLists = mainCtx.elements.some(el => el.type === 'list' || el.type === 'richList')
      const allImages = [...mainCtx.images, ...headerCtx.images, ...footerCtx.images]
      const imageExts = allImages.map(img => img.ext)
      const headerImageOffset = mainCtx.images.length
      const footerImageOffset = mainCtx.images.length + headerCtx.images.length

      const files: ZipEntry[] = [
        { name: '[Content_Types].xml', data: enc.encode(generateDocxContentTypes(hasLists, hasHeader, hasFooter, imageExts)) },
        { name: '_rels/.rels', data: enc.encode(DOCX_RELS) },
        { name: 'word/_rels/document.xml.rels', data: enc.encode(generateDocxRels(hasLists, hasHeader, hasFooter, mainCtx.hyperlinks, mainCtx.images, 0)) },
        { name: 'word/document.xml', data: enc.encode(generateDocxDocument(mainCtx.elements, hasHeader ? 'rIdHeader' : undefined, hasFooter ? 'rIdFooter' : undefined)) },
        { name: 'word/styles.xml', data: enc.encode(generateDocxStyles()) },
        ...(hasLists ? [{ name: 'word/numbering.xml', data: enc.encode(generateDocxNumbering()) }] : []),
        ...(hasHeader ? [
          { name: 'word/header1.xml', data: enc.encode(generateDocxHeader(headerCtx.elements, headerCtx.hyperlinks.length > 0 || headerCtx.images.length > 0)) },
          ...(headerCtx.hyperlinks.length > 0 || headerCtx.images.length > 0 ? [{ name: 'word/_rels/header1.xml.rels', data: enc.encode(generatePartRels(headerCtx.hyperlinks, headerCtx.images, headerImageOffset)) }] : [])
        ] : []),
        ...(hasFooter ? [
          { name: 'word/footer1.xml', data: enc.encode(generateDocxFooter(footerCtx.elements, footerCtx.hyperlinks.length > 0 || footerCtx.images.length > 0)) },
          ...(footerCtx.hyperlinks.length > 0 || footerCtx.images.length > 0 ? [{ name: 'word/_rels/footer1.xml.rels', data: enc.encode(generatePartRels(footerCtx.hyperlinks, footerCtx.images, footerImageOffset)) }] : [])
        ] : []),
        ...allImages.map((img, i) => ({ name: `word/media/image${i + 1}.${img.ext}`, data: img.data }))
      ]
      return createZip(files)
    }
  }
  return builder
}

const ODT_MIMETYPE = 'application/vnd.oasis.opendocument.text'

const generateOdtManifest = (images: { ext: string }[]): string => {
  const imgEntries = images.map((img, i) => `<manifest:file-entry manifest:full-path="Pictures/image${i + 1}.${img.ext}" manifest:media-type="${img.ext === 'jpeg' ? 'image/jpeg' : img.ext === 'gif' ? 'image/gif' : 'image/png'}"/>`).join('')
  return `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="${NS.manifest}" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>${imgEntries}</manifest:manifest>`
}

type OdtStyleAcc = { styles: string[]; counter: number }

const buildOdtTextProps = (opts?: TextOptions, size?: number): string => {
  const font = opts?.code ? COURIER : opts?.font
  return [
    font ? ` style:font-name="${font}"` : '',
    size || opts?.size ? ` fo:font-size="${size || opts?.size}pt"` : '',
    opts?.bold ? ' fo:font-weight="bold"' : '',
    opts?.italic ? ' fo:font-style="italic"' : '',
    opts?.underline ? ' style:text-underline-style="solid"' : '',
    opts?.strikethrough ? ' style:text-line-through-style="solid"' : '',
    opts?.code ? ' fo:background-color="#E8E8E8"' : '',
    opts?.color ? ` fo:color="${opts.color}"` : ''
  ].join('')
}

const buildOdtStyle = (name: string, family: 'paragraph' | 'text', opts?: TextOptions, size?: number): string => {
  const textProps = buildOdtTextProps(opts, size)
  const pProps = family === 'paragraph' && opts?.align ? `<style:paragraph-properties fo:text-align="${opts.align}"/>` : ''
  return `<style:style style:name="${name}" style:family="${family}">${pProps}<style:text-properties${textProps}/></style:style>`
}

const needsTextStyle = (opts?: TextOptions): boolean =>
  !!(opts?.bold || opts?.italic || opts?.code || opts?.underline || opts?.strikethrough || opts?.color)

const buildOdtRuns = (runs: TextRun[], acc: OdtStyleAcc): string =>
  runs.map(run => {
    if (needsTextStyle(run.opts)) {
      const styleName = `T${++acc.counter}`
      acc.styles.push(buildOdtStyle(styleName, 'text', run.opts))
      return `<text:span text:style-name="${styleName}">${escapeXml(run.text)}</text:span>`
    }
    return escapeXml(run.text)
  }).join('')

const buildOdtList = (items: ListItem[], ordered: boolean, acc: OdtStyleAcc): string => {
  const listItems = items.map(item =>
    `<text:list-item><text:p text:style-name="Standard">${buildOdtRuns(item.runs, acc)}</text:p>${item.children ? buildOdtList(item.children.items, item.children.ordered, acc) : ''}</text:list-item>`
  ).join('')
  return `<text:list text:style-name="${ordered ? 'Numbering_20_1' : 'List_20_1'}">${listItems}</text:list>`
}

const elementToOdt = (el: DocElement, acc: OdtStyleAcc, imageOffset: number): string => {
  switch (el.type) {
    case 'heading':
      return `<text:h text:style-name="Heading${el.level}" text:outline-level="${el.level}">${escapeXml(el.text)}</text:h>`
    case 'paragraph': {
      if (el.opts?.bold || el.opts?.italic || el.opts?.color || el.opts?.align || el.opts?.font || el.opts?.underline) {
        const sn = `P${++acc.counter}`
        acc.styles.push(buildOdtStyle(sn, 'paragraph', el.opts))
        return `<text:p text:style-name="${sn}">${escapeXml(el.text)}</text:p>`
      }
      return `<text:p text:style-name="Standard">${escapeXml(el.text)}</text:p>`
    }
    case 'text': {
      const sn = `P${++acc.counter}`
      acc.styles.push(buildOdtStyle(sn, 'paragraph', el.opts, el.size))
      return `<text:p text:style-name="${sn}">${escapeXml(el.text)}</text:p>`
    }
    case 'lineBreak': return '<text:p text:style-name="Standard"/>'
    case 'horizontalRule': return '<text:p text:style-name="Standard">────────────────────────────────────────</text:p>'
    case 'list':
      return `<text:list text:style-name="${el.ordered ? 'Numbering_20_1' : 'List_20_1'}">${el.items.map(item => `<text:list-item><text:p text:style-name="Standard">${escapeXml(item)}</text:p></text:list-item>`).join('')}</text:list>`
    case 'table':
      return `<table:table xmlns:table="${NS.table}">${el.rows.map(row => `<table:table-row>${row.map(cell => `<table:table-cell><text:p>${escapeXml(cell)}</text:p></table:table-cell>`).join('')}</table:table-row>`).join('')}</table:table>`
    case 'link':
      return `<text:p><text:a xlink:href="${escapeXml(el.url)}" xmlns:xlink="${NS.xlink}">${escapeXml(el.text)}</text:a></text:p>`
    case 'image': {
      const idx = imageOffset + 1
      const wCm = (el.opts.width * 2.54).toFixed(3), hCm = (el.opts.height * 2.54).toFixed(3)
      return `<text:p><draw:frame draw:name="Image${idx}" svg:width="${wCm}cm" svg:height="${hCm}cm" xmlns:draw="${NS.draw}" xmlns:svg="${NS.svg}"><draw:image xlink:href="Pictures/image${idx}.${detectImageType(el.data)}" xmlns:xlink="${NS.xlink}"/></draw:frame></text:p>`
    }
    case 'pageNumber':
      return '<text:p><text:page-number text:select-page="current"/></text:p>'
    case 'richParagraph':
      return `<text:p text:style-name="Standard">${buildOdtRuns(el.runs, acc)}</text:p>`
    case 'richList': return buildOdtList(el.items, el.ordered, acc)
    case 'richTable':
      return `<table:table xmlns:table="${NS.table}">${el.rows.map(row => `<table:table-row>${row.map(cell => `<table:table-cell><text:p>${buildOdtRuns(cell, acc)}</text:p></table:table-cell>`).join('')}</table:table-row>`).join('')}</table:table>`
    case 'blockquote': return el.elements.map(child => elementToOdt(child, acc, imageOffset)).join('')
    case 'codeBlock': {
      const sn = `P${++acc.counter}`
      acc.styles.push(`<style:style style:name="${sn}" style:family="paragraph"><style:text-properties style:font-name="${COURIER}" fo:background-color="#F5F5F5"/></style:style>`)
      return el.code.split('\n').map(line => `<text:p text:style-name="${sn}">${escapeXml(line)}</text:p>`).join('')
    }
  }
}

const generateOdtContent = (elements: DocElement[], imageOffset = 0): string => {
  const acc: OdtStyleAcc = { styles: [], counter: 0 }
  const body = elements.map(el => elementToOdt(el, acc, imageOffset)).join('')
  const autoStyles = acc.styles.length > 0 ? `<office:automatic-styles>${acc.styles.join('')}</office:automatic-styles>` : ''
  return `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content xmlns:office="${NS.office}" xmlns:text="${NS.text}" xmlns:style="${NS.style}" xmlns:fo="${NS.fo}" office:version="1.2">\n${autoStyles}\n<office:body>\n<office:text>\n${body}\n</office:text>\n</office:body>\n</office:document-content>`
}

const generateOdtStyles = (headerElements?: DocElement[], footerElements?: DocElement[]): string => {
  const headings = [1, 2, 3, 4, 5, 6].map(i => {
    const sizes = { 1: '24pt', 2: '18pt', 3: '14pt', 4: '12pt', 5: '10pt', 6: '9pt' } as Record<number, string>
    return `<style:style style:name="Heading${i}" style:family="paragraph"><style:text-properties fo:font-size="${sizes[i]}" fo:font-weight="bold"/></style:style>`
  }).join('')
  const hasMaster = headerElements || footerElements
  let masterStyles = ''
  if (hasMaster) {
    const acc: OdtStyleAcc = { styles: [], counter: 1000 }
    const hdr = headerElements ? `<style:header>${headerElements.map(el => elementToOdt(el, acc, 0)).join('')}</style:header>` : ''
    const ftr = footerElements ? `<style:footer>${footerElements.map(el => elementToOdt(el, acc, 0)).join('')}</style:footer>` : ''
    masterStyles = `<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1">${hdr}${ftr}</style:master-page></office:master-styles>`
  }
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="${NS.office}" xmlns:style="${NS.style}" xmlns:fo="${NS.fo}" xmlns:text="${NS.text}" office:version="1.2"><office:styles><style:style style:name="Standard" style:family="paragraph"/>${headings}</office:styles>${masterStyles}</office:document-styles>`
}

/**
 * Create a new ODT document
 */
export function odt(): ODTBuilder {
  const docPrIdCounter = { value: 1 }
  const ctx: BuildContext = { elements: [], hyperlinks: [], images: [], nextRId: 1, nextDocPrId: docPrIdCounter }
  const headerCtx: BuildContext = { elements: [], hyperlinks: [], images: [], nextRId: 100, nextDocPrId: docPrIdCounter }
  const footerCtx: BuildContext = { elements: [], hyperlinks: [], images: [], nextRId: 200, nextDocPrId: docPrIdCounter }
  let hasHeader = false, hasFooter = false
  const builder: ODTBuilder = {
    content(fn) { fn(createContext(ctx)); return builder },
    header(fn) { hasHeader = true; fn(createContext(headerCtx)); return builder },
    footer(fn) { hasFooter = true; fn(createContext(footerCtx)); return builder },
    build() {
      const enc = new TextEncoder()
      const allImages = [...ctx.images, ...headerCtx.images, ...footerCtx.images]
      return createZip([
        { name: 'mimetype', data: enc.encode(ODT_MIMETYPE) },
        { name: 'META-INF/manifest.xml', data: enc.encode(generateOdtManifest(allImages)) },
        { name: 'content.xml', data: enc.encode(generateOdtContent(ctx.elements)) },
        { name: 'styles.xml', data: enc.encode(generateOdtStyles(hasHeader ? headerCtx.elements : undefined, hasFooter ? footerCtx.elements : undefined)) },
        ...allImages.map((img, i) => ({ name: `Pictures/image${i + 1}.${img.ext}`, data: img.data }))
      ])
    }
  }
  return builder
}

const parseMarkdownToElements = (md: string): DocElement[] => astToElements(parseMarkdownAST(md))

export function markdownToDocx(md: string): Uint8Array {
  const docPrIdCounter = { value: 1 }
  const mainCtx: BuildContext = { elements: [], hyperlinks: [], images: [], nextRId: 10, nextDocPrId: docPrIdCounter }
  mainCtx.elements.push(...parseMarkdownToElements(md))
  const enc = new TextEncoder()
  const hasLists = mainCtx.elements.some(el => el.type === 'list' || el.type === 'richList')
  const files: ZipEntry[] = [
    { name: '[Content_Types].xml', data: enc.encode(generateDocxContentTypes(hasLists, false, false, [])) },
    { name: '_rels/.rels', data: enc.encode(DOCX_RELS) },
    { name: 'word/_rels/document.xml.rels', data: enc.encode(generateDocxRels(hasLists, false, false, [], [], 0)) },
    { name: 'word/document.xml', data: enc.encode(generateDocxDocument(mainCtx.elements, undefined, undefined)) },
    { name: 'word/styles.xml', data: enc.encode(generateDocxStyles()) },
    ...(hasLists ? [{ name: 'word/numbering.xml', data: enc.encode(generateDocxNumbering()) }] : [])
  ]
  return createZip(files)
}

export function markdownToOdt(md: string): Uint8Array {
  const enc = new TextEncoder()
  return createZip([
    { name: 'mimetype', data: enc.encode(ODT_MIMETYPE) },
    { name: 'META-INF/manifest.xml', data: enc.encode(generateOdtManifest([])) },
    { name: 'content.xml', data: enc.encode(generateOdtContent(parseMarkdownToElements(md))) },
    { name: 'styles.xml', data: enc.encode(generateOdtStyles()) }
  ])
}

export default docx
