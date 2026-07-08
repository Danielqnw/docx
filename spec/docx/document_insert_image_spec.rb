# frozen_string_literal: true

require 'spec_helper'
require 'docx'
require 'zip'
require 'stringio'

RSpec.describe Docx::Document, '#insert_image_at_placeholder' do
  XML_NS = Docx::Document::XML_NAMESPACES

  def fake_png(w, h)
    "\x89PNG\r\n\x1A\n".b + "\x00\x00\x00\rIHDR".b + [w, h].pack('N2') + "\x08\x02\x00\x00\x00".b
  end

  def build_no_image_fixture(path)
    TableFixtureBuilder.build(path) do |b|
      b.add_paragraph('{{photo}}')
      spec = TableFixtureBuilder::TableSpec.new(rows: 1, cols: 1)
      spec.set_text(0, 0, '{{img}}')
      b.add_table(spec)
    end
  end

  def open_fixture
    path = File.join(Dir.mktmpdir, 'no_image.docx')
    build_no_image_fixture(path)
    Docx::Document.open(path)
  end

  def stream_and_reopen(doc)
    stream = doc.stream
    stream.rewind
    Docx::Document.open(stream)
  end

  def body_paragraph_with_drawing(doc)
    doc.doc.at_xpath('//w:body/w:p[.//w:drawing]', XML_NS)
  end

  def table_cell_with_drawing(doc)
    doc.doc.at_xpath('//w:tc[.//w:drawing]', XML_NS)
  end

  describe 'insert_image_at_placeholder' do
    it 'inserts a drawing in a body paragraph, registers rels and content type, and cleans placeholder' do
      doc = open_fixture
      png = fake_png(200, 100)

      result = doc.insert_image_at_placeholder('{{photo}}', StringIO.new(png))

      expect(result[:relationship_id]).to match(/\ArId\d+\z/)
      expect(result[:entry_path]).to match(%r{\Aword/media/image_generated_\d+\.png\z})
      expect(result[:fit]).to eq(:contain)

      paragraph = body_paragraph_with_drawing(doc)
      expect(paragraph).not_to be_nil
      expect(paragraph.xpath('.//w:drawing', XML_NS)).not_to be_empty

      blip = paragraph.at_xpath('.//a:blip/@r:embed', XML_NS)
      expect(blip.value).to eq(result[:relationship_id])

      expect(doc.images).to include(result[:relationship_id] => result[:entry_path])

      paragraph_text = paragraph.xpath('.//w:t', XML_NS).map(&:text).join
      expect(paragraph_text).not_to include('{{photo}}')

      reopened = stream_and_reopen(doc)
      expect(reopened.doc.at_xpath('//w:drawing', XML_NS)).not_to be_nil
    end

    it 'inserts a drawing in a table cell placeholder' do
      doc = open_fixture
      png = fake_png(200, 100)

      result = doc.insert_image_at_placeholder('{{img}}', StringIO.new(png))

      cell = table_cell_with_drawing(doc)
      expect(cell).not_to be_nil
      expect(cell.xpath('.//w:drawing', XML_NS)).not_to be_empty

      blip = cell.at_xpath('.//a:blip/@r:embed', XML_NS)
      expect(blip.value).to eq(result[:relationship_id])
      expect(doc.images).to include(result[:relationship_id] => result[:entry_path])

      cell_text = cell.xpath('.//w:t', XML_NS).map(&:text).join
      expect(cell_text).not_to include('{{img}}')

      reopened = stream_and_reopen(doc)
      expect(reopened.doc.at_xpath('//w:tc//w:drawing', XML_NS)).not_to be_nil
    end

    it 'keeps placeholder text when cleanup_placeholder is false' do
      doc = open_fixture
      png = fake_png(200, 100)

      doc.insert_image_at_placeholder('{{photo}}', StringIO.new(png), cleanup_placeholder: false)

      paragraph = doc.doc.xpath('//w:p', XML_NS).find do |p|
        p.xpath('.//w:t', XML_NS).map(&:text).join.include?('{{photo}}')
      end
      paragraph_text = paragraph.xpath('.//w:t', XML_NS).map(&:text).join
      expect(paragraph_text).to include('{{photo}}')
    end

    it 'raises when placeholder cannot be found' do
      doc = open_fixture
      png = fake_png(200, 100)

      expect { doc.insert_image_at_placeholder('{{missing}}', StringIO.new(png)) }
        .to raise_error(Docx::Errors::ImagePlaceholderNotFound)
    end

    it 'writes [Content_Types].xml with png Default on stream' do
      doc = open_fixture
      png = fake_png(200, 100)

      doc.insert_image_at_placeholder('{{photo}}', StringIO.new(png))
      out_path = File.join(Dir.mktmpdir, 'out.docx')
      doc.save(out_path)

      content_types_xml = Zip::File.open(out_path) { |zip| zip.read('[Content_Types].xml') }
      expect(content_types_xml).to include('Extension="png"')
      expect(content_types_xml).to include('ContentType="image/png"')
    end
  end

  describe 'insert_images_at_placeholder' do
    it 'inserts multiple drawings with correct relationship count' do
      doc = open_fixture
      png1 = fake_png(200, 100)
      png2 = fake_png(300, 150)

      results = doc.insert_images_at_placeholder(
        '{{photo}}',
        [StringIO.new(png1), StringIO.new(png2)]
      )

      expect(results.length).to eq(2)
      expect(results.map { |r| r[:relationship_id] }).to all(match(/\ArId\d+\z/))
      expect(results.map { |r| r[:relationship_id] }.uniq.length).to eq(2)

      paragraph = body_paragraph_with_drawing(doc)
      drawings = paragraph.xpath('.//w:drawing', XML_NS)
      expect(drawings.length).to eq(2)

      image_rels = doc.image_relationships
      expect(image_rels.length).to eq(2)

      paragraph_text = paragraph.xpath('.//w:t', XML_NS).map(&:text).join
      expect(paragraph_text).not_to include('{{photo}}')
    end
  end
end
