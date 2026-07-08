# frozen_string_literal: true

require 'spec_helper'
require 'docx'
require 'tmpdir'

describe 'Docx::Document#substitute_across_runs' do
  it 'replaces placeholders in body paragraphs and table cells' do
    Dir.mktmpdir do |dir|
      path = File.join(dir, 'across.docx')
      TableFixtureBuilder.build(path) do |b|
        b.add_paragraph('Hello {{name}}')
        spec = TableFixtureBuilder::TableSpec.new(rows: 1, cols: 1)
        spec.set_text(0, 0, 'City: {{city}}')
        b.add_table(spec)
      end

      doc = Docx::Document.open(path)
      doc.substitute_across_runs('{{name}}', 'Alice')
      doc.substitute_across_runs('{{city}}', 'Paris')

      reopened = Docx::Document.open(doc.stream)
      full_text = reopened.paragraphs.map(&:text).join("\n") +
                  reopened.tables.flat_map { |t| t.rows.flat_map { |r| r.cells.map(&:text) } }.join("\n")

      expect(full_text).to include('Hello Alice')
      expect(full_text).to include('Paris')
      expect(full_text).not_to include('{{name}}')
      expect(full_text).not_to include('{{city}}')
    end
  end

  it 'matches a placeholder split across multiple runs' do
    doc_xml = <<~XML
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r><w:rPr><w:b/></w:rPr><w:t>{{ti</w:t></w:r>
            <w:r><w:t>tle}}</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>
    XML

    rels_xml = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'

    doc = Docx::Document.allocate
    doc.instance_variable_set(:@doc, Nokogiri::XML(doc_xml))
    doc.instance_variable_set(:@rels, Nokogiri::XML(rels_xml))

    doc.substitute_across_runs('{{title}}', 'Report')

    text = doc.instance_variable_get(:@doc).xpath('//w:t').map(&:text).join
    expect(text).to eq('Report')
    # formatting of the first matched run is preserved
    expect(doc.instance_variable_get(:@doc).at_xpath('//w:r[1]//w:b')).not_to be_nil
  end
end
