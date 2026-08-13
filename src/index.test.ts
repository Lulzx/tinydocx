import { describe, expect, test } from 'bun:test'
import { inflateRawSync } from 'node:zlib'
import { docx, odt, markdownToDocx, markdownToOdt } from './index'

/** Extract all file contents from a ZIP as a concatenated string */
const readZip = (zip: Uint8Array): string => {
  const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength)
  const parts: string[] = []
  const dec = new TextDecoder()
  let pos = 0
  while (pos + 4 <= zip.length && view.getUint32(pos, true) === 0x04034b50) {
    const method = view.getUint16(pos + 8, true)
    const compressedSize = view.getUint32(pos + 18, true)
    const nameLen = view.getUint16(pos + 26, true)
    const extraLen = view.getUint16(pos + 28, true)
    const name = dec.decode(zip.subarray(pos + 30, pos + 30 + nameLen))
    const dataStart = pos + 30 + nameLen + extraLen
    const raw = zip.subarray(dataStart, dataStart + compressedSize)
    parts.push(name)
    try {
      parts.push(method === 8 ? dec.decode(inflateRawSync(raw)) : dec.decode(raw))
    } catch { parts.push(dec.decode(raw)) }
    pos = dataStart + compressedSize
  }
  return parts.join('\n')
}

/** Extract a single ZIP entry by exact name. */
const readZipEntry = (zip: Uint8Array, wanted: string): string => {
  const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength)
  const dec = new TextDecoder()
  let pos = 0
  while (pos + 30 <= zip.length && view.getUint32(pos, true) === 0x04034b50) {
    const method = view.getUint16(pos + 8, true)
    const compressedSize = view.getUint32(pos + 18, true)
    const nameLen = view.getUint16(pos + 26, true)
    const extraLen = view.getUint16(pos + 28, true)
    const name = dec.decode(zip.subarray(pos + 30, pos + 30 + nameLen))
    const dataStart = pos + 30 + nameLen + extraLen
    const raw = zip.subarray(dataStart, dataStart + compressedSize)
    if (name === wanted) return dec.decode(method === 8 ? inflateRawSync(raw) : raw)
    pos = dataStart + compressedSize
  }
  throw new Error(`ZIP entry not found: ${wanted}`)
}

const firstZipMethod = (zip: Uint8Array): number =>
  new DataView(zip.buffer, zip.byteOffset, zip.byteLength).getUint16(8, true)

describe('docx', () => {
  test('creates valid ZIP archive', () => {
    const doc = docx()
    doc.content(() => {})
    const bytes = doc.build()
    expect(bytes[0]).toBe(0x50)
    expect(bytes[1]).toBe(0x4b)
  })

  test('returns Uint8Array', () => {
    const doc = docx()
    doc.content(() => {})
    expect(doc.build()).toBeInstanceOf(Uint8Array)
  })

  test('includes required DOCX files', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Hello'))
    const str = readZip(doc.build())
    expect(str).toContain('[Content_Types].xml')
    expect(str).toContain('word/document.xml')
    expect(str).toContain('_rels/.rels')
    expect(str).toContain('word/styles.xml')
  })

  test('renders heading level 1', () => {
    const doc = docx()
    doc.content((ctx) => ctx.heading('Test Heading', 1))
    const str = readZip(doc.build())
    expect(str).toContain('Test Heading')
    expect(str).toContain('Heading1')
    expect(str).toContain('w:sz w:val="48"')
  })

  test('renders heading level 2', () => {
    const doc = docx()
    doc.content((ctx) => ctx.heading('H2 Heading', 2))
    const str = readZip(doc.build())
    expect(str).toContain('H2 Heading')
    expect(str).toContain('Heading2')
    expect(str).toContain('w:sz w:val="36"')
  })

  test('renders heading level 3', () => {
    const doc = docx()
    doc.content((ctx) => ctx.heading('H3 Heading', 3))
    const str = readZip(doc.build())
    expect(str).toContain('H3 Heading')
    expect(str).toContain('Heading3')
    expect(str).toContain('w:sz w:val="28"')
  })

  test('renders paragraph', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Test paragraph'))
    const str = readZip(doc.build())
    expect(str).toContain('Test paragraph')
  })

  test('renders text with size', () => {
    const doc = docx()
    doc.content((ctx) => ctx.text('Large text', 24))
    const str = readZip(doc.build())
    expect(str).toContain('Large text')
    expect(str).toContain('w:sz w:val="48"')
  })

  test('renders text with size and options', () => {
    const doc = docx()
    doc.content((ctx) => ctx.text('Styled text', 16, { bold: true, color: '#0000ff' }))
    const str = readZip(doc.build())
    expect(str).toContain('Styled text')
    expect(str).toContain('w:sz w:val="32"')
    expect(str).toContain('<w:b/>')
    expect(str).toContain('w:color w:val="0000ff"')
  })

  test('applies bold formatting', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Bold text', { bold: true }))
    const str = readZip(doc.build())
    expect(str).toContain('<w:b/>')
  })

  test('applies italic formatting', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Italic text', { italic: true }))
    const str = readZip(doc.build())
    expect(str).toContain('<w:i/>')
  })

  test('applies underline formatting', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Underlined', { underline: true }))
    const str = readZip(doc.build())
    expect(str).toContain('w:u w:val="single"')
  })

  test('applies combined formatting', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('All styles', { bold: true, italic: true, underline: true }))
    const str = readZip(doc.build())
    expect(str).toContain('<w:b/>')
    expect(str).toContain('<w:i/>')
    expect(str).toContain('w:u w:val="single"')
  })

  test('applies color', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Red text', { color: '#ff0000' }))
    const str = readZip(doc.build())
    expect(str).toContain('w:color w:val="ff0000"')
  })

  test('applies left alignment', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Left', { align: 'left' }))
    const str = readZip(doc.build())
    expect(str).toContain('w:jc w:val="left"')
  })

  test('applies center alignment', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Centered', { align: 'center' }))
    const str = readZip(doc.build())
    expect(str).toContain('w:jc w:val="center"')
  })

  test('applies right alignment', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Right', { align: 'right' }))
    const str = readZip(doc.build())
    expect(str).toContain('w:jc w:val="right"')
  })

  test('applies custom font', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Arial text', { font: 'Arial' }))
    const str = readZip(doc.build())
    expect(str).toContain('w:rFonts w:ascii="Arial" w:hAnsi="Arial"')
  })

  test('applies size option in TextOptions', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Sized text', { size: 20 }))
    const str = readZip(doc.build())
    expect(str).toContain('w:sz w:val="40"')
  })

  test('renders line break', () => {
    const doc = docx()
    doc.content((ctx) => {
      ctx.paragraph('Before')
      ctx.lineBreak()
      ctx.paragraph('After')
    })
    const str = readZip(doc.build())
    expect(str).toContain('Before')
    expect(str).toContain('After')
    expect(str).toContain('<w:p/>')
  })

  test('renders horizontal rule', () => {
    const doc = docx()
    doc.content((ctx) => ctx.horizontalRule())
    const str = readZip(doc.build())
    expect(str).toContain('w:pBdr')
    expect(str).toContain('w:bottom')
  })

  test('escapes XML special characters', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Test <tag> & "quotes" \'apostrophe\''))
    const str = readZip(doc.build())
    expect(str).toContain('&lt;tag&gt;')
    expect(str).toContain('&amp;')
    expect(str).toContain('&quot;')
    expect(str).toContain('&apos;')
  })

  test('supports method chaining', () => {
    const doc = docx()
    const result = doc.content((ctx) => ctx.paragraph('Test'))
    expect(result).toBe(doc)
  })

  test('supports header chaining', () => {
    const doc = docx()
    const result = doc.header((ctx) => ctx.paragraph('Header'))
    expect(result).toBe(doc)
  })

  test('supports footer chaining', () => {
    const doc = docx()
    const result = doc.footer((ctx) => ctx.paragraph('Footer'))
    expect(result).toBe(doc)
  })

  test('renders bullet list', () => {
    const doc = docx()
    doc.content((ctx) => ctx.list(['Item 1', 'Item 2', 'Item 3']))
    const str = readZip(doc.build())
    expect(str).toContain('Item 1')
    expect(str).toContain('Item 2')
    expect(str).toContain('Item 3')
    expect(str).toContain('w:numId w:val="1"')
    expect(str).toContain('word/numbering.xml')
  })

  test('renders numbered list', () => {
    const doc = docx()
    doc.content((ctx) => ctx.list(['First', 'Second'], true))
    const str = readZip(doc.build())
    expect(str).toContain('First')
    expect(str).toContain('Second')
    expect(str).toContain('w:numId w:val="2"')
  })

  test('renders multiple lists', () => {
    const doc = docx()
    doc.content((ctx) => {
      ctx.list(['A', 'B'])
      ctx.list(['1', '2'], true)
    })
    const str = readZip(doc.build())
    expect(str).toContain('A')
    expect(str).toContain('B')
    expect(str).toContain('1')
    expect(str).toContain('2')
  })

  test('renders single item list', () => {
    const doc = docx()
    doc.content((ctx) => ctx.list(['Only one']))
    const str = readZip(doc.build())
    expect(str).toContain('Only one')
  })

  test('renders table', () => {
    const doc = docx()
    doc.content((ctx) => ctx.table([
      ['A', 'B'],
      ['C', 'D']
    ]))
    const str = readZip(doc.build())
    expect(str).toContain('<w:tbl>')
    expect(str).toContain('<w:tr>')
    expect(str).toContain('<w:tc>')
    expect(str).toContain('A')
    expect(str).toContain('B')
    expect(str).toContain('C')
    expect(str).toContain('D')
  })

  test('renders table with column widths', () => {
    const doc = docx()
    doc.content((ctx) => ctx.table([['X', 'Y']], { colWidths: [2000, 3000] }))
    const str = readZip(doc.build())
    expect(str).toContain('w:gridCol w:w="2000"')
    expect(str).toContain('w:gridCol w:w="3000"')
    expect(str).toContain('w:tcW w:w="2000"')
    expect(str).toContain('w:tcW w:w="3000"')
  })

  test('renders table with borders', () => {
    const doc = docx()
    doc.content((ctx) => ctx.table([['Cell']]))
    const str = readZip(doc.build())
    expect(str).toContain('w:tblBorders')
    expect(str).toContain('w:top')
    expect(str).toContain('w:bottom')
    expect(str).toContain('w:left')
    expect(str).toContain('w:right')
    expect(str).toContain('w:insideH')
    expect(str).toContain('w:insideV')
  })

  test('renders large table', () => {
    const rows = Array.from({ length: 10 }, (_, i) =>
      Array.from({ length: 5 }, (_, j) => `R${i}C${j}`)
    )
    const doc = docx()
    doc.content((ctx) => ctx.table(rows))
    const str = readZip(doc.build())
    expect(str).toContain('R0C0')
    expect(str).toContain('R9C4')
  })

  test('renders hyperlink', () => {
    const doc = docx()
    doc.content((ctx) => ctx.link('Click here', 'https://example.com'))
    const str = readZip(doc.build())
    expect(str).toContain('Click here')
    expect(str).toContain('w:hyperlink')
    expect(str).toContain('https://example.com')
    expect(str).toContain('TargetMode="External"')
  })

  test('renders hyperlink with styling', () => {
    const doc = docx()
    doc.content((ctx) => ctx.link('Styled link', 'https://example.com', { bold: true }))
    const str = readZip(doc.build())
    expect(str).toContain('Styled link')
    expect(str).toContain('<w:b/>')
    expect(str).toContain('w:u w:val="single"')
  })

  test('renders multiple hyperlinks', () => {
    const doc = docx()
    doc.content((ctx) => {
      ctx.link('Link 1', 'https://one.com')
      ctx.link('Link 2', 'https://two.com')
    })
    const str = readZip(doc.build())
    expect(str).toContain('Link 1')
    expect(str).toContain('Link 2')
    expect(str).toContain('https://one.com')
    expect(str).toContain('https://two.com')
  })

  test('renders PNG image', () => {
    const pngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    const doc = docx()
    doc.content((ctx) => ctx.image(pngBytes, { width: 2, height: 1 }))
    const str = readZip(doc.build())
    expect(str).toContain('w:drawing')
    expect(str).toContain('wp:inline')
    expect(str).toContain('word/media/image1.png')
    expect(str).toContain('image/png')
  })

  test('detects JPEG image type', () => {
    const jpegBytes = new Uint8Array([0xff, 0xd8, 0xff, 0xe0])
    const doc = docx()
    doc.content((ctx) => ctx.image(jpegBytes, { width: 1, height: 1 }))
    const str = readZip(doc.build())
    expect(str).toContain('image/jpeg')
    expect(str).toContain('image1.jpeg')
  })

  test('detects GIF image type', () => {
    const gifBytes = new Uint8Array([0x47, 0x49, 0x46, 0x38])
    const doc = docx()
    doc.content((ctx) => ctx.image(gifBytes, { width: 1, height: 1 }))
    const str = readZip(doc.build())
    expect(str).toContain('image/gif')
    expect(str).toContain('image1.gif')
  })

  test('detects WebP image type', () => {
    const webpBytes = new Uint8Array([0x52, 0x49, 0x46, 0x46, 0x00, 0x00, 0x00, 0x00, 0x57, 0x45, 0x42, 0x50])
    const doc = docx()
    doc.content((ctx) => ctx.image(webpBytes, { width: 1, height: 1 }))
    const str = readZip(doc.build())
    expect(str).toContain('image1.webp')
    expect(str).toContain('image/webp')
  })

  test('renders multiple images', () => {
    const pngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47])
    const jpegBytes = new Uint8Array([0xff, 0xd8, 0xff, 0xe0])
    const doc = docx()
    doc.content((ctx) => {
      ctx.image(pngBytes, { width: 1, height: 1 })
      ctx.image(jpegBytes, { width: 2, height: 2 })
    })
    const str = readZip(doc.build())
    expect(str).toContain('image1.png')
    expect(str).toContain('image2.jpeg')
  })

  test('assigns unique docPr IDs to multiple images', () => {
    const pngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47])
    const doc = docx()
    doc.content((ctx) => {
      ctx.image(pngBytes, { width: 1, height: 1 })
      ctx.image(pngBytes, { width: 1, height: 1 })
      ctx.image(pngBytes, { width: 1, height: 1 })
    })
    const str = readZip(doc.build())
    expect(str).toContain('wp:docPr id="1"')
    expect(str).toContain('wp:docPr id="2"')
    expect(str).toContain('wp:docPr id="3"')
  })

  test('calculates image dimensions correctly', () => {
    const pngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47])
    const doc = docx()
    doc.content((ctx) => ctx.image(pngBytes, { width: 2, height: 1.5 }))
    const str = readZip(doc.build())
    expect(str).toContain('cx="1828800"')
    expect(str).toContain('cy="1371600"')
  })

  test('renders header', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Body'))
    doc.header((ctx) => ctx.paragraph('Header text'))
    const str = readZip(doc.build())
    expect(str).toContain('word/header1.xml')
    expect(str).toContain('Header text')
    expect(str).toContain('w:headerReference')
  })

  test('renders footer', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Body'))
    doc.footer((ctx) => ctx.paragraph('Footer text'))
    const str = readZip(doc.build())
    expect(str).toContain('word/footer1.xml')
    expect(str).toContain('Footer text')
    expect(str).toContain('w:footerReference')
  })

  test('renders header and footer together', () => {
    const doc = docx()
    doc.header((ctx) => ctx.paragraph('Header'))
    doc.footer((ctx) => ctx.paragraph('Footer'))
    doc.content((ctx) => ctx.paragraph('Body'))
    const str = readZip(doc.build())
    expect(str).toContain('word/header1.xml')
    expect(str).toContain('word/footer1.xml')
    expect(str).toContain('Header')
    expect(str).toContain('Footer')
    expect(str).toContain('Body')
  })

  test('header with hyperlink creates separate rels file', () => {
    const doc = docx()
    doc.header((ctx) => ctx.link('Click', 'https://example.com'))
    doc.content((ctx) => ctx.paragraph('Body'))
    const str = readZip(doc.build())
    expect(str).toContain('word/_rels/header1.xml.rels')
    expect(str).toContain('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
  })

  test('footer with image creates separate rels file', () => {
    const pngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47])
    const doc = docx()
    doc.footer((ctx) => ctx.image(pngBytes, { width: 1, height: 1 }))
    doc.content((ctx) => ctx.paragraph('Body'))
    const str = readZip(doc.build())
    expect(str).toContain('word/_rels/footer1.xml.rels')
  })

  test('images in header use correct media index', () => {
    const pngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47])
    const doc = docx()
    doc.content((ctx) => ctx.image(pngBytes, { width: 1, height: 1 }))
    doc.header((ctx) => ctx.image(pngBytes, { width: 1, height: 1 }))
    const str = readZip(doc.build())
    expect(str).toContain('word/media/image1.png')
    expect(str).toContain('word/media/image2.png')
  })

  test('images across content, header and footer share unique docPr IDs', () => {
    const pngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47])
    const doc = docx()
    doc.content((ctx) => ctx.image(pngBytes, { width: 1, height: 1 }))
    doc.header((ctx) => ctx.image(pngBytes, { width: 1, height: 1 }))
    doc.footer((ctx) => ctx.image(pngBytes, { width: 1, height: 1 }))
    const str = readZip(doc.build())
    expect(str).toContain('wp:docPr id="1"')
    expect(str).toContain('wp:docPr id="2"')
    expect(str).toContain('wp:docPr id="3"')
  })

  test('renders page number', () => {
    const doc = docx()
    doc.footer((ctx) => ctx.pageNumber())
    doc.content((ctx) => ctx.paragraph('Body'))
    const str = readZip(doc.build())
    expect(str).toContain('w:fldChar')
    expect(str).toContain('PAGE')
    expect(str).toContain('w:instrText')
    expect(str).toContain('fldCharType="begin"')
    expect(str).toContain('fldCharType="separate"')
    expect(str).toContain('fldCharType="end"')
  })

  test('includes styles.xml with defaults', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Test'))
    const str = readZip(doc.build())
    expect(str).toContain('word/styles.xml')
    expect(str).toContain('w:docDefaults')
    expect(str).toContain('Calibri')
  })

  test('includes heading styles in styles.xml', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Test'))
    const str = readZip(doc.build())
    expect(str).toContain('Heading 1')
    expect(str).toContain('Heading 2')
    expect(str).toContain('Heading 3')
  })

  test('excludes numbering.xml when no lists', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('No lists'))
    const str = readZip(doc.build())
    expect(str).not.toContain('word/numbering.xml')
  })

  test('includes numbering.xml when lists present', () => {
    const doc = docx()
    doc.content((ctx) => ctx.list(['Item']))
    const str = readZip(doc.build())
    expect(str).toContain('word/numbering.xml')
    expect(str).toContain('w:abstractNum')
    expect(str).toContain('w:num')
  })

  test('complex document with all features', () => {
    const pngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47])
    const doc = docx()
    doc.header((ctx) => ctx.paragraph('Company', { bold: true }))
    doc.footer((ctx) => ctx.pageNumber())
    doc.content((ctx) => {
      ctx.heading('Title', 1)
      ctx.paragraph('Intro', { italic: true })
      ctx.list(['A', 'B', 'C'])
      ctx.table([['X', 'Y'], ['1', '2']])
      ctx.link('Website', 'https://example.com')
      ctx.image(pngBytes, { width: 1, height: 1 })
      ctx.horizontalRule()
      ctx.paragraph('End', { align: 'center' })
    })
    const bytes = doc.build()
    expect(bytes.length).toBeGreaterThan(100)
    const str = readZip(bytes)
    expect(str).toContain('Title')
    expect(str).toContain('Intro')
    expect(str).toContain('Website')
    expect(str).toContain('word/header1.xml')
    expect(str).toContain('word/footer1.xml')
  })

  test('produces compressed output (smaller than uncompressed)', () => {
    const doc = docx()
    doc.content((ctx) => {
      for (let i = 0; i < 100; i++) ctx.paragraph(`Paragraph number ${i} with some repeated text content`)
    })
    const bytes = doc.build()
    const str = readZip(bytes)
    expect(str).toContain('Paragraph number 0')
    expect(str).toContain('Paragraph number 99')
    // With DEFLATE the output should be significantly smaller
    expect(bytes.length).toBeLessThan(str.length)
  })

  test('multi-level numbering includes 5 levels', () => {
    const doc = docx()
    doc.content((ctx) => ctx.list(['Item']))
    const str = readZip(doc.build())
    expect(str).toContain('w:ilvl="0"')
    expect(str).toContain('w:ilvl="1"')
    expect(str).toContain('w:ilvl="2"')
    expect(str).toContain('w:ilvl="3"')
    expect(str).toContain('w:ilvl="4"')
    expect(str).toContain('w:ind w:left="720"')
    expect(str).toContain('w:ind w:left="1440"')
    expect(str).toContain('w:ind w:left="2160"')
  })

  test('bullet levels use alternating characters', () => {
    const doc = docx()
    doc.content((ctx) => ctx.list(['Item']))
    const str = readZip(doc.build())
    expect(str).toContain('w:lvlText w:val="•"')
    expect(str).toContain('w:lvlText w:val="◦"')
    expect(str).toContain('w:lvlText w:val="▪"')
  })
})

describe('odt', () => {
  test('creates valid ZIP archive', () => {
    const doc = odt()
    doc.content(() => {})
    const bytes = doc.build()
    expect(bytes[0]).toBe(0x50)
    expect(bytes[1]).toBe(0x4b)
  })

  test('returns Uint8Array', () => {
    const doc = odt()
    doc.content(() => {})
    expect(doc.build()).toBeInstanceOf(Uint8Array)
  })

  test('includes mimetype as first file', () => {
    const doc = odt()
    doc.content(() => {})
    const str = readZip(doc.build())
    expect(str).toContain('mimetype')
    expect(str).toContain('application/vnd.oasis.opendocument.text')
    expect(firstZipMethod(doc.build())).toBe(0)
  })

  test('includes required ODT files', () => {
    const doc = odt()
    doc.content((ctx) => ctx.paragraph('Hello'))
    const str = readZip(doc.build())
    expect(str).toContain('content.xml')
    expect(str).toContain('styles.xml')
    expect(str).toContain('META-INF/manifest.xml')
  })

  test('renders heading level 1', () => {
    const doc = odt()
    doc.content((ctx) => ctx.heading('Test Heading', 1))
    const str = readZip(doc.build())
    expect(str).toContain('Test Heading')
    expect(str).toContain('Heading1')
    expect(str).toContain('text:outline-level="1"')
  })

  test('renders heading level 2', () => {
    const doc = odt()
    doc.content((ctx) => ctx.heading('H2', 2))
    const str = readZip(doc.build())
    expect(str).toContain('Heading2')
    expect(str).toContain('text:outline-level="2"')
  })

  test('renders heading level 3', () => {
    const doc = odt()
    doc.content((ctx) => ctx.heading('H3', 3))
    const str = readZip(doc.build())
    expect(str).toContain('Heading3')
    expect(str).toContain('text:outline-level="3"')
  })

  test('renders paragraph', () => {
    const doc = odt()
    doc.content((ctx) => ctx.paragraph('Test paragraph'))
    const str = readZip(doc.build())
    expect(str).toContain('Test paragraph')
  })

  test('renders text with size', () => {
    const doc = odt()
    doc.content((ctx) => ctx.text('Large text', 24))
    const str = readZip(doc.build())
    expect(str).toContain('Large text')
    expect(str).toContain('fo:font-size="24pt"')
  })

  test('applies bold formatting', () => {
    const doc = odt()
    doc.content((ctx) => ctx.paragraph('Bold text', { bold: true }))
    const str = readZip(doc.build())
    expect(str).toContain('fo:font-weight="bold"')
  })

  test('applies italic formatting', () => {
    const doc = odt()
    doc.content((ctx) => ctx.paragraph('Italic text', { italic: true }))
    const str = readZip(doc.build())
    expect(str).toContain('fo:font-style="italic"')
  })

  test('applies underline formatting', () => {
    const doc = odt()
    doc.content((ctx) => ctx.paragraph('Underlined', { underline: true }))
    const str = readZip(doc.build())
    expect(str).toContain('style:text-underline-style="solid"')
  })

  test('applies combined formatting', () => {
    const doc = odt()
    doc.content((ctx) => ctx.paragraph('All styles', { bold: true, italic: true, underline: true }))
    const str = readZip(doc.build())
    expect(str).toContain('fo:font-weight="bold"')
    expect(str).toContain('fo:font-style="italic"')
    expect(str).toContain('style:text-underline-style="solid"')
  })

  test('applies color', () => {
    const doc = odt()
    doc.content((ctx) => ctx.paragraph('Red text', { color: '#ff0000' }))
    const str = readZip(doc.build())
    expect(str).toContain('fo:color="#ff0000"')
  })

  test('applies alignment', () => {
    const doc = odt()
    doc.content((ctx) => ctx.paragraph('Centered', { align: 'center' }))
    const str = readZip(doc.build())
    expect(str).toContain('fo:text-align="center"')
  })

  test('applies custom font', () => {
    const doc = odt()
    doc.content((ctx) => ctx.paragraph('Arial text', { font: 'Arial' }))
    const str = readZip(doc.build())
    expect(str).toContain('style:font-name="Arial"')
  })

  test('renders horizontal rule', () => {
    const doc = odt()
    doc.content((ctx) => ctx.horizontalRule())
    const str = readZip(doc.build())
    expect(str).toContain('────')
  })

  test('escapes XML special characters', () => {
    const doc = odt()
    doc.content((ctx) => ctx.paragraph('Test <tag> & "quotes"'))
    const str = readZip(doc.build())
    expect(str).toContain('&lt;tag&gt;')
    expect(str).toContain('&amp;')
    expect(str).toContain('&quot;')
  })

  test('supports method chaining', () => {
    const doc = odt()
    const result = doc.content((ctx) => ctx.paragraph('Test'))
    expect(result).toBe(doc)
  })

  test('renders bullet list', () => {
    const doc = odt()
    doc.content((ctx) => ctx.list(['Item 1', 'Item 2']))
    const str = readZip(doc.build())
    expect(str).toContain('text:list')
    expect(str).toContain('text:list-item')
    expect(str).toContain('Item 1')
    expect(str).toContain('Item 2')
  })

  test('renders numbered list', () => {
    const doc = odt()
    doc.content((ctx) => ctx.list(['First', 'Second'], true))
    const str = readZip(doc.build())
    expect(str).toContain('text:list')
    expect(str).toContain('Numbering')
  })

  test('renders table', () => {
    const doc = odt()
    doc.content((ctx) => ctx.table([
      ['A', 'B'],
      ['C', 'D']
    ]))
    const str = readZip(doc.build())
    expect(str).toContain('table:table')
    expect(str).toContain('table:table-row')
    expect(str).toContain('table:table-cell')
    expect(str).toContain('A')
    expect(str).toContain('D')
  })

  test('renders hyperlink', () => {
    const doc = odt()
    doc.content((ctx) => ctx.link('Click here', 'https://example.com'))
    const str = readZip(doc.build())
    expect(str).toContain('text:a')
    expect(str).toContain('xlink:href')
    expect(str).toContain('https://example.com')
  })

  test('creates automatic styles for formatted text', () => {
    const doc = odt()
    doc.content((ctx) => {
      ctx.paragraph('Normal')
      ctx.paragraph('Bold', { bold: true })
      ctx.paragraph('Italic', { italic: true })
    })
    const str = readZip(doc.build())
    expect(str).toContain('office:automatic-styles')
    expect(str).toContain('style:style')
  })

  test('renders image with draw:frame', () => {
    const pngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47])
    const doc = odt()
    doc.content((ctx) => ctx.image(pngBytes, { width: 2, height: 1 }))
    const str = readZip(doc.build())
    expect(str).toContain('draw:frame')
    expect(str).toContain('draw:image')
    expect(str).toContain('Pictures/image1.png')
    expect(str).toContain('svg:width')
    expect(str).toContain('svg:height')
  })

  test('image stored in Pictures/ and manifest updated', () => {
    const pngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47])
    const doc = odt()
    doc.content((ctx) => ctx.image(pngBytes, { width: 1, height: 1 }))
    const str = readZip(doc.build())
    expect(str).toContain('Pictures/image1.png')
    expect(str).toContain('manifest:full-path="Pictures/image1.png"')
  })

  test('multiple images reference their matching files', () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47])
    const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xe0])
    const doc = odt().content(ctx => {
      ctx.image(png, { width: 1, height: 1 })
      ctx.image(jpeg, { width: 1, height: 1 })
    })
    const content = readZipEntry(doc.build(), 'content.xml')
    expect(content).toContain('Pictures/image1.png')
    expect(content).toContain('Pictures/image2.jpeg')
  })

  test('declares WebP with the correct media type', () => {
    const webp = new Uint8Array([0x52, 0x49, 0x46, 0x46, 0, 0, 0, 0, 0x57, 0x45, 0x42, 0x50])
    const manifest = readZipEntry(odt().content(ctx => ctx.image(webp, { width: 1, height: 1 })).build(), 'META-INF/manifest.xml')
    expect(manifest).toContain('manifest:media-type="image/webp"')
  })

  test('defines referenced list styles', () => {
    const zip = odt().content(ctx => ctx.list(['Bullet'])).build()
    const content = readZipEntry(zip, 'content.xml')
    const styles = readZipEntry(zip, 'styles.xml')
    expect(content).toContain('text:style-name="List_20_1"')
    expect(styles).toContain('style:name="List_20_1"')
    expect(styles).toContain('style:name="Numbering_20_1"')
  })

  test('renders page number', () => {
    const doc = odt()
    doc.content((ctx) => ctx.pageNumber())
    const str = readZip(doc.build())
    expect(str).toContain('text:page-number')
    expect(str).toContain('text:select-page="current"')
  })

  test('renders header and footer', () => {
    const doc = odt()
    doc.header((ctx) => ctx.paragraph('Header Text'))
    doc.footer((ctx) => ctx.paragraph('Footer Text'))
    doc.content((ctx) => ctx.paragraph('Body'))
    const str = readZip(doc.build())
    expect(str).toContain('style:header')
    expect(str).toContain('Header Text')
    expect(str).toContain('style:footer')
    expect(str).toContain('Footer Text')
    expect(str).toContain('style:master-page')
  })

  test('defines styles referenced by formatted headers and footers', () => {
    const zip = odt()
      .header(ctx => ctx.paragraph('Bold header', { bold: true }))
      .footer(ctx => ctx.paragraph('Italic footer', { italic: true }))
      .build()
    const styles = readZipEntry(zip, 'styles.xml')
    expect(styles).toContain('text:style-name="P1001"')
    expect(styles).toContain('style:name="P1001"')
    expect(styles).toContain('fo:font-weight="bold"')
    expect(styles).toContain('fo:font-style="italic"')
  })

  test('supports header chaining', () => {
    const doc = odt()
    const result = doc.header((ctx) => ctx.paragraph('Header'))
    expect(result).toBe(doc)
  })

  test('supports footer chaining', () => {
    const doc = odt()
    const result = doc.footer((ctx) => ctx.paragraph('Footer'))
    expect(result).toBe(doc)
  })

  test('footer with page number in ODT', () => {
    const doc = odt()
    doc.footer((ctx) => ctx.pageNumber())
    doc.content((ctx) => ctx.paragraph('Body'))
    const str = readZip(doc.build())
    expect(str).toContain('text:page-number')
    expect(str).toContain('style:footer')
  })
})

describe('markdownToDocx', () => {
  test('returns Uint8Array', () => {
    const bytes = markdownToDocx('# Hello')
    expect(bytes).toBeInstanceOf(Uint8Array)
  })

  test('converts h1 header', () => {
    const str = readZip(markdownToDocx('# Heading 1'))
    expect(str).toContain('Heading 1')
    expect(str).toContain('Heading1')
  })

  test('converts h2 header', () => {
    const str = readZip(markdownToDocx('## Heading 2'))
    expect(str).toContain('Heading 2')
    expect(str).toContain('Heading2')
  })

  test('converts h3 header', () => {
    const str = readZip(markdownToDocx('### Heading 3'))
    expect(str).toContain('Heading 3')
    expect(str).toContain('Heading3')
  })

  test('converts bullet lists with dash', () => {
    const str = readZip(markdownToDocx('- Item 1\n- Item 2'))
    expect(str).toContain('Item 1')
    expect(str).toContain('Item 2')
    expect(str).toContain('w:numId w:val="1"')
  })

  test('converts bullet lists with asterisk', () => {
    const str = readZip(markdownToDocx('* Item A\n* Item B'))
    expect(str).toContain('Item A')
    expect(str).toContain('Item B')
  })

  test('converts numbered lists', () => {
    const str = readZip(markdownToDocx('1. First\n2. Second\n3. Third'))
    expect(str).toContain('First')
    expect(str).toContain('Second')
    expect(str).toContain('Third')
    expect(str).toContain('w:numId w:val="2"')
  })

  test('converts horizontal rules with dashes', () => {
    const str = readZip(markdownToDocx('---'))
    expect(str).toContain('w:pBdr')
  })

  test('converts horizontal rules with asterisks', () => {
    const str = readZip(markdownToDocx('***'))
    expect(str).toContain('w:pBdr')
  })

  test('converts horizontal rules with underscores', () => {
    const str = readZip(markdownToDocx('___'))
    expect(str).toContain('w:pBdr')
  })

  test('converts paragraphs', () => {
    const str = readZip(markdownToDocx('This is a paragraph.\n\nThis is another.'))
    expect(str).toContain('This is a paragraph.')
    expect(str).toContain('This is another.')
  })

  test('handles empty input', () => {
    const bytes = markdownToDocx('')
    expect(bytes).toBeInstanceOf(Uint8Array)
    expect(bytes.length).toBeGreaterThan(0)
  })

  test('handles complex markdown', () => {
    const md = `# Title

Introduction paragraph.

## Section 1

- Point A
- Point B

## Section 2

1. First step
2. Second step

---

Conclusion.`
    const str = readZip(markdownToDocx(md))
    expect(str).toContain('Title')
    expect(str).toContain('Section 1')
    expect(str).toContain('Section 2')
    expect(str).toContain('Point A')
    expect(str).toContain('First step')
    expect(str).toContain('Conclusion')
  })

  test('multi-line paragraphs joined with space', () => {
    const str = readZip(markdownToDocx('Line one\nline two\nline three'))
    expect(str).toContain('Line one line two line three')
  })

  test('multi-line paragraphs stop at blank line', () => {
    const str = readZip(markdownToDocx('First line\nsecond line\n\nNew paragraph'))
    expect(str).toContain('First line second line')
    expect(str).toContain('New paragraph')
  })

  test('multi-line paragraphs stop at special lines', () => {
    const str = readZip(markdownToDocx('Some text\nmore text\n# Heading'))
    expect(str).toContain('Some text more text')
    expect(str).toContain('Heading')
  })

  test('richList detected for numbering.xml inclusion', () => {
    const str = readZip(markdownToDocx('- Item A\n- Item B'))
    expect(str).toContain('word/numbering.xml')
  })

  test('converts inline links with a relationship', () => {
    const zip = markdownToDocx('Visit [Example](https://example.com).')
    const document = readZipEntry(zip, 'word/document.xml')
    const rels = readZipEntry(zip, 'word/_rels/document.xml.rels')
    expect(document).toContain('<w:hyperlink r:id="rId10">')
    expect(rels).toContain('Id="rId10"')
    expect(rels).toContain('Target="https://example.com"')
  })

  test('preserves inline links in headings', () => {
    const zip = markdownToDocx('# [Linked heading](https://example.com/heading)')
    const document = readZipEntry(zip, 'word/document.xml')
    const rels = readZipEntry(zip, 'word/_rels/document.xml.rels')
    expect(document).toContain('<w:hyperlink r:id="rId10">')
    expect(rels).toContain('Target="https://example.com/heading"')
  })

  test('converts Markdown images as external image relationships', () => {
    const zip = markdownToDocx('![Logo](https://example.com/logo.png)')
    const document = readZipEntry(zip, 'word/document.xml')
    const rels = readZipEntry(zip, 'word/_rels/document.xml.rels')
    expect(document).toContain('r:link="rId10"')
    expect(document).toContain('name="Logo"')
    expect(rels).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"')
    expect(rels).toContain('Target="https://example.com/logo.png" TargetMode="External"')
  })
})

describe('markdownToOdt', () => {
  test('returns Uint8Array', () => {
    const bytes = markdownToOdt('# Hello')
    expect(bytes).toBeInstanceOf(Uint8Array)
  })

  test('converts headers', () => {
    const str = readZip(markdownToOdt('# H1\n## H2\n### H3'))
    expect(str).toContain('H1')
    expect(str).toContain('H2')
    expect(str).toContain('H3')
  })

  test('converts bullet lists', () => {
    const str = readZip(markdownToOdt('- Item 1\n- Item 2'))
    expect(str).toContain('Item 1')
    expect(str).toContain('Item 2')
    expect(str).toContain('text:list')
  })

  test('converts numbered lists', () => {
    const str = readZip(markdownToOdt('1. First\n2. Second'))
    expect(str).toContain('First')
    expect(str).toContain('Second')
  })

  test('handles empty input', () => {
    const bytes = markdownToOdt('')
    expect(bytes).toBeInstanceOf(Uint8Array)
    expect(bytes.length).toBeGreaterThan(0)
  })

  test('converts inline links and images', () => {
    const content = readZipEntry(markdownToOdt('[Example](https://example.com) ![Logo](https://example.com/logo.png)'), 'content.xml')
    expect(content).toContain('<text:a xlink:href="https://example.com">Example</text:a>')
    expect(content).toContain('<draw:image xlink:href="https://example.com/logo.png"/>')
  })

  test('preserves inline links in headings', () => {
    const content = readZipEntry(markdownToOdt('# [Linked heading](https://example.com/heading)'), 'content.xml')
    expect(content).toContain('<text:h')
    expect(content).toContain('<text:a xlink:href="https://example.com/heading">Linked heading</text:a>')
  })
})
