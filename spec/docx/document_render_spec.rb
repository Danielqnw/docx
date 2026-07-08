# frozen_string_literal: true

require 'spec_helper'
require 'docx'
require 'tmpdir'
require 'zip'
require 'stringio'

RSpec.describe Docx::Document, '#render' do
  XML_NS = Docx::Document::XML_NAMESPACES

  def fake_png(w, h)
    "\x89PNG\r\n\x1A\n".b + "\x00\x00\x00\rIHDR".b + [w, h].pack('N2') + "\x08\x02\x00\x00\x00".b
  end

  def build_comprehensive_fixture(path, include_leftover: false)
    TableFixtureBuilder.build(path) do |b|
      b.add_paragraph('Name: {{name}}')
      b.add_paragraph('Bio: {{bio}}')
      b.add_paragraph('{{photo}}')
      b.add_paragraph("\u2610 选项A")
      b.add_paragraph('Left: {{leftover}}') if include_leftover

      spec = TableFixtureBuilder::TableSpec.new(rows: 2, cols: 2)
      spec.set_text(0, 0, 'Header1')
      spec.set_text(0, 1, 'Header2')
      spec.set_text(1, 0, '{{row}} {{c1}}')
      spec.set_text(1, 1, '{{c2}}')
      b.add_table(spec)
    end
  end

  def reopen(doc)
    Docx::Document.open(doc.stream)
  end

  def table_data_rows(doc)
    table = doc.tables.first
    table.rows.drop(1)
  end

  def cell_texts(row)
    row.cells.map(&:text)
  end

  describe '#render' do
    it 'replaces text placeholders including multiline bio with w:br' do
      Dir.mktmpdir do |dir|
        path = File.join(dir, 'render_text.docx')
        build_comprehensive_fixture(path)

        doc = Docx::Document.open(path)
        doc.render(
          text: {
            '{{name}}' => 'Alice',
            '{{bio}}' => "line1\nline2"
          },
          multiline: true
        )

        reopened = reopen(doc)
        body_text = reopened.paragraphs.map(&:text).join("\n")

        expect(body_text).to include('Name: Alice')
        expect(body_text).not_to include('{{name}}')

        bio_paragraph = reopened.doc.xpath('//w:p[contains(.//w:t/text(), "Bio:")]', XML_NS).first
        expect(bio_paragraph.xpath('.//w:br', XML_NS)).not_to be_empty
        bio_text = bio_paragraph.xpath('.//w:t', XML_NS).map(&:text).join
        expect(bio_text).to include('line1')
        expect(bio_text).to include('line2')
      end
    end

    it 'duplicates template table rows and removes the template row' do
      Dir.mktmpdir do |dir|
        path = File.join(dir, 'render_table.docx')
        build_comprehensive_fixture(path)

        doc = Docx::Document.open(path)
        doc.render(
          tables: [
            {
              placeholder_row: '{{row}}',
              rows: [
                { '{{c1}}' => 'a1', '{{c2}}' => 'a2' },
                { '{{c1}}' => 'b1', '{{c2}}' => 'b2' }
              ]
            }
          ]
        )

        reopened = reopen(doc)
        data_rows = table_data_rows(reopened)

        expect(data_rows.length).to eq(2)
        expect(cell_texts(data_rows[0]).map(&:strip)).to eq(['a1', 'a2'])
        expect(cell_texts(data_rows[1]).map(&:strip)).to eq(['b1', 'b2'])

        full_table_text = reopened.tables.first.rows.flat_map { |r| r.cells.map(&:text) }.join
        expect(full_table_text).not_to include('{{row}}')
        expect(full_table_text).not_to include('{{c1}}')
        expect(full_table_text).not_to include('{{c2}}')
      end
    end

    it 'inserts an image at a placeholder without a pre-existing image' do
      Dir.mktmpdir do |dir|
        path = File.join(dir, 'render_image.docx')
        build_comprehensive_fixture(path)

        doc = Docx::Document.open(path)
        png = fake_png(200, 100)
        doc.render(images: { '{{photo}}' => StringIO.new(png) })

        expect(doc.doc.at_xpath('//w:drawing', XML_NS)).not_to be_nil
        expect(doc.images.values).to include(a_string_matching(%r{\Aword/media/image_generated_\d+\.png\z}))

        paragraph = doc.doc.at_xpath('//w:p[.//w:drawing]', XML_NS)
        paragraph_text = paragraph.xpath('.//w:t', XML_NS).map(&:text).join
        expect(paragraph_text).not_to include('{{photo}}')

        content_types = doc.instance_variable_get(:@replace)['[Content_Types].xml']
        expect(content_types).to include('image/png')

        reopened = reopen(doc)
        expect(reopened.doc.at_xpath('//w:drawing', XML_NS)).not_to be_nil
      end
    end

    it 'checks character-based checkboxes' do
      Dir.mktmpdir do |dir|
        path = File.join(dir, 'render_checkbox.docx')
        build_comprehensive_fixture(path)

        doc = Docx::Document.open(path)
        doc.render(checkboxes: { '选项A' => true })

        checkbox_text = doc.doc.xpath('//w:p', XML_NS)[3].xpath('.//w:t', XML_NS).map(&:text).join
        expect(checkbox_text).to eq("\u2611 选项A")

        reopened = reopen(doc)
        reopened_text = reopened.doc.xpath('//w:p', XML_NS)[3].xpath('.//w:t', XML_NS).map(&:text).join
        expect(reopened_text).to eq("\u2611 选项A")
      end
    end

    it 'strips unfilled placeholders when strip_unfilled is true' do
      Dir.mktmpdir do |dir|
        path = File.join(dir, 'render_strip.docx')
        build_comprehensive_fixture(path, include_leftover: true)

        doc = Docx::Document.open(path)
        doc.render(
          text: { '{{name}}' => 'Alice' },
          strip_unfilled: true
        )

        body_text = doc.paragraphs.map(&:text).join("\n")
        expect(body_text).to include('Name: Alice')
        expect(body_text).not_to include('{{leftover}}')
        expect(body_text).to include('Left: ')
        expect(body_text).not_to include('{{name}}')
        expect(body_text).not_to include('{{bio}}')
        expect(body_text).not_to include('{{photo}}')

        reopened = reopen(doc)
        reopened_text = reopened.paragraphs.map(&:text).join("\n")
        expect(reopened_text).not_to include('{{leftover}}')
        expect(reopened_text).not_to include('{{bio}}')
        expect(reopened_text).not_to include('{{photo}}')
      end
    end

    it 'skips missing image placeholders when strict is false' do
      Dir.mktmpdir do |dir|
        path = File.join(dir, 'render_non_strict.docx')
        build_comprehensive_fixture(path)

        doc = Docx::Document.open(path)
        expect do
          doc.render(
            text: { '{{name}}' => 'Bob' },
            images: { '{{missing_photo}}' => StringIO.new(fake_png(10, 10)) },
            strict: false
          )
        end.not_to raise_error

        reopened = reopen(doc)
        expect(reopened.paragraphs.first.text).to include('Bob')
      end
    end

    it 'performs a comprehensive end-to-end render' do
      Dir.mktmpdir do |dir|
        path = File.join(dir, 'render_e2e.docx')
        build_comprehensive_fixture(path)

        doc = Docx::Document.open(path)
        doc.render(
          text: {
            '{{name}}' => 'Alice',
            '{{bio}}' => "alpha\nbeta"
          },
          tables: [
            {
              placeholder_row: '{{row}}',
              rows: [
                { '{{c1}}' => 'x1', '{{c2}}' => 'y1' },
                { '{{c1}}' => 'x2', '{{c2}}' => 'y2' }
              ]
            }
          ],
          images: { '{{photo}}' => StringIO.new(fake_png(120, 80)) },
          checkboxes: { '选项A' => true },
          strip_unfilled: true,
          multiline: true
        )

        reopened = reopen(doc)

        expect(reopened.paragraphs.map(&:text).join).to include('Name: Alice')
        expect(reopened.doc.at_xpath('//w:p[contains(.//w:t/text(), "Bio:")]//w:br', XML_NS)).not_to be_nil
        expect(reopened.doc.at_xpath('//w:drawing', XML_NS)).not_to be_nil

        data_rows = table_data_rows(reopened)
        expect(data_rows.length).to eq(2)
        expect(cell_texts(data_rows[0]).map(&:strip)).to eq(['x1', 'y1'])
        expect(cell_texts(data_rows[1]).map(&:strip)).to eq(['x2', 'y2'])

        checkbox_text = reopened.doc.xpath('//w:p', XML_NS)[3].xpath('.//w:t', XML_NS).map(&:text).join
        expect(checkbox_text).to eq("\u2611 选项A")

        full_text = reopened.to_s + reopened.tables.flat_map { |t| t.rows.flat_map { |r| r.cells.map(&:text) } }.join
        expect(full_text).not_to include('{{')
      end
    end
  end

  describe '#strip_unfilled_placeholders' do
    it 'clears placeholders split across multiple runs' do
      doc_xml = <<~XML
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r><w:t>before {{x</w:t></w:r>
              <w:r><w:t>}} after</w:t></w:r>
            </w:p>
            <w:p><w:r><w:t>filled: Alice</w:t></w:r></w:p>
          </w:body>
        </w:document>
      XML

      rels_xml = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'

      doc = Docx::Document.allocate
      doc.instance_variable_set(:@doc, Nokogiri::XML(doc_xml))
      doc.instance_variable_set(:@rels, Nokogiri::XML(rels_xml))

      doc.strip_unfilled_placeholders

      first_paragraph_text = doc.instance_variable_get(:@doc).xpath('//w:p[1]//w:t').map(&:text).join
      expect(first_paragraph_text).to eq('before  after')
      expect(first_paragraph_text).not_to include('{{')

      second_paragraph_text = doc.instance_variable_get(:@doc).xpath('//w:p[2]//w:t').map(&:text).join
      expect(second_paragraph_text).to eq('filled: Alice')

      doc.strip_unfilled_placeholders
      expect(doc.instance_variable_get(:@doc).xpath('//w:p[1]//w:t').map(&:text).join).to eq('before  after')
    end
  end
end
