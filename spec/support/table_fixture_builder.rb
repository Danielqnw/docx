# frozen_string_literal: true

require 'fileutils'
require 'set'
require 'zip'

module TableFixtureBuilder
  W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
  CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

  CellDef = Struct.new(:text, :colspan, :vmerge, keyword_init: true)

  class TableSpec
    attr_reader :row_count, :col_count

    def initialize(rows:, cols:)
      @row_count = rows
      @col_count = cols
      @cells = {}
      @skipped = Set.new
    end

    def default_cell_text(row, col)
      "r#{row}c#{col}"
    end

    def set_text(row, col, text)
      cell_def(row, col).text = text
    end

    def fill_default_text!
      (0...@row_count).each do |row|
        (0...@col_count).each do |col|
          next if @skipped.include?([row, col])

          cell = cell_def(row, col)
          next if cell.vmerge == :continue

          cell.text ||= default_cell_text(row, col)
        end
      end
    end

    def merge_horizontal(row, col_start, col_end)
      validate_col_range!(col_start, col_end)

      colspan = col_end - col_start + 1
      cell_def(row, col_start).colspan = colspan

      ((col_start + 1)..col_end).each do |col|
        @skipped << [row, col]
      end
    end

    def merge_vertical(col, row_start, row_end)
      validate_row_range!(row_start, row_end)

      cell_def(row_start, col).vmerge = :restart

      ((row_start + 1)..row_end).each do |row|
        cell_def(row, col).vmerge = :continue
      end
    end

    def merge_rect(row0, col0, row1, col1)
      validate_row_range!(row0, row1)
      validate_col_range!(col0, col1)

      colspan = col1 - col0 + 1
      cell_def(row0, col0).colspan = colspan
      cell_def(row0, col0).vmerge = :restart

      (row0..row1).each do |row|
        ((col0 + 1)..col1).each { |col| @skipped << [row, col] }
      end

      ((row0 + 1)..row1).each do |row|
        cell_def(row, col0).colspan = colspan
        cell_def(row, col0).vmerge = :continue
      end
    end

    def cell_definition_at(row, col)
      return nil if @skipped.include?([row, col])

      cell_def(row, col)
    end

    private

    def cell_def(row, col)
      @cells[[row, col]] ||= CellDef.new(text: nil, colspan: 1, vmerge: nil)
    end

    def validate_row_range!(row_start, row_end)
      raise ArgumentError, "row range out of bounds" if row_start.negative? || row_end >= @row_count || row_start > row_end
    end

    def validate_col_range!(col_start, col_end)
      raise ArgumentError, "column range out of bounds" if col_start.negative? || col_end >= @col_count || col_start > col_end
    end
  end

  class DocumentBuilder
    def initialize
      @paragraphs = []
      @tables = []
    end

    def add_paragraph(text)
      @paragraphs << text
      self
    end

    def add_table(spec)
      @tables << spec
      self
    end

    def write(path)
      FileUtils.mkdir_p(File.dirname(path)) unless File.dirname(path) == '.'

      Zip::OutputStream.open(path) do |out|
        package_entries.each do |name, content|
          out.put_next_entry(name)
          out.write(content)
        end
      end

      path
    end

    private

    def package_entries
      {
        '[Content_Types].xml' => content_types_xml,
        '_rels/.rels' => root_rels_xml,
        'word/_rels/document.xml.rels' => document_rels_xml,
        'word/styles.xml' => styles_xml,
        'word/document.xml' => document_xml
      }
    end

    def content_types_xml
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="#{CONTENT_TYPES_NS}">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
        </Types>
      XML
    end

    def root_rels_xml
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="#{PACKAGE_REL_NS}">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
      XML
    end

    def document_rels_xml
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="#{PACKAGE_REL_NS}">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        </Relationships>
      XML
    end

    def styles_xml
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="#{W_NS}">
          <w:docDefaults>
            <w:rPrDefault>
              <w:rPr>
                <w:sz w:val="22"/>
              </w:rPr>
            </w:rPrDefault>
            <w:pPrDefault/>
          </w:docDefaults>
          <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
            <w:name w:val="Normal"/>
          </w:style>
        </w:styles>
      XML
    end

    def document_xml
      body = @paragraphs.map { |text| paragraph_xml(text) }.join
      body += @tables.map { |spec| table_xml(spec) }.join

      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="#{W_NS}">
          <w:body>
            #{body}
            <w:sectPr>
              <w:pgSz w:w="12240" w:h="15840"/>
              <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
            </w:sectPr>
          </w:body>
        </w:document>
      XML
    end

    def paragraph_xml(text)
      <<~XML
        <w:p>
          <w:r>
            <w:t>#{escape_xml(text)}</w:t>
          </w:r>
        </w:p>
      XML
    end

    def table_xml(spec)
      spec.fill_default_text!
      grid_cols = (0...spec.col_count).map { '<w:gridCol w:w="2400"/>' }.join
      rows = (0...spec.row_count).map { |row| table_row_xml(spec, row) }.join

      <<~XML
        <w:tbl>
          <w:tblGrid>
            #{grid_cols}
          </w:tblGrid>
          #{rows}
        </w:tbl>
      XML
    end

    def table_row_xml(spec, row)
      cells = []
      col = 0

      while col < spec.col_count
        if spec.cell_definition_at(row, col).nil?
          col += 1
          next
        end

        cell = spec.cell_definition_at(row, col)
        cells << table_cell_xml(cell)
        col += cell.colspan || 1
      end

      "<w:tr>#{cells.join}</w:tr>"
    end

    def table_cell_xml(cell)
      tc_pr = tc_pr_xml(cell)
      text = cell.text || ''

      <<~XML
        <w:tc>
          #{tc_pr}
          <w:p>
            <w:r>
              <w:t>#{escape_xml(text)}</w:t>
            </w:r>
          </w:p>
        </w:tc>
      XML
    end

    def tc_pr_xml(cell)
      parts = []

      colspan = cell.colspan || 1
      parts << %(<w:gridSpan w:val="#{colspan}"/>) if colspan > 1

      case cell.vmerge
      when :restart
        parts << '<w:vMerge w:val="restart"/>'
      when :continue
        parts << '<w:vMerge/>'
      end

      return '<w:tcPr/>' if parts.empty?

      "<w:tcPr>#{parts.join}</w:tcPr>"
    end

    def escape_xml(text)
      text.to_s
          .gsub('&', '&amp;')
          .gsub('<', '&lt;')
          .gsub('>', '&gt;')
          .gsub('"', '&quot;')
          .gsub("'", '&apos;')
    end
  end

  module_function

  def build(path, &block)
    builder = DocumentBuilder.new
    yield(builder) if block
    builder.write(path)
  end

  def plain_3x3(path)
    spec = TableSpec.new(rows: 3, cols: 3)
    build(path) { |doc| doc.add_table(spec) }
  end

  def horizontal_merge(path, row: 0, col_start: 0, col_end: 1, rows: 3, cols: 3)
    spec = TableSpec.new(rows: rows, cols: cols)
    spec.merge_horizontal(row, col_start, col_end)
    build(path) { |doc| doc.add_table(spec) }
  end

  def vertical_merge(path, col: 0, row_start: 0, row_end: 2, rows: 3, cols: 3)
    spec = TableSpec.new(rows: rows, cols: cols)
    spec.merge_vertical(col, row_start, row_end)
    build(path) { |doc| doc.add_table(spec) }
  end

  def rect_merge(path, row0: 0, col0: 0, row1: 1, col1: 1, rows: 3, cols: 3)
    spec = TableSpec.new(rows: rows, cols: cols)
    spec.merge_rect(row0, col0, row1, col1)
    build(path) { |doc| doc.add_table(spec) }
  end

  FIXTURES = {
    'plain_3x3.docx' => :plain_3x3,
    'horizontal_merge.docx' => :horizontal_merge,
    'vertical_merge.docx' => :vertical_merge,
    'rect_merge.docx' => :rect_merge
  }.freeze

  def write_all_fixtures(output_dir)
    FileUtils.mkdir_p(output_dir)

    FIXTURES.each do |filename, builder_method|
      path = File.join(output_dir, filename)
      public_send(builder_method, path)
    end
  end
end
