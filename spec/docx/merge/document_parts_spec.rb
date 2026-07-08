# frozen_string_literal: true

require 'spec_helper'
require 'docx'
require 'securerandom'
require 'tmpdir'
require 'zip'

describe 'Docx::Document part persistence APIs' do
  fixtures_path = 'spec/fixtures'
  content_types_ns = Docx::Document::CONTENT_TYPES_NS
  rels_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'

  def save_to_tempfile(doc)
    path = File.join(Dir.tmpdir, "docx_merge_#{SecureRandom.hex(8)}.docx")
    doc.save(path)
    path
  end

  def content_types_xml(path)
    Zip::File.open(path) { |zip| zip.read(Docx::Document::CONTENT_TYPES_PATH) }
  end

  def rels_xml(path)
    Zip::File.open(path) do |zip|
      entry = zip.glob('word/_rels/document*.xml.rels').first
      zip.read(entry.name)
    end
  end

  describe '#add_part' do
    it 'persists a new zip entry and registers a Content_Types Override' do
      doc = Docx::Document.open(File.join(fixtures_path, 'basic.docx'))
      payload = '<item>custom</item>'
      zip_path = 'customXml/item1.xml'
      content_type = 'application/xml'

      expect(doc.add_part(zip_path, content_type, payload)).to eq(zip_path)

      temp_path = save_to_tempfile(doc)
      reopened = Docx::Document.open(temp_path)

      entry_bytes = nil
      Zip::File.open(temp_path) { |zip| entry_bytes = zip.read(zip_path) }
      expect(entry_bytes).to eq(payload)

      types = Nokogiri::XML(content_types_xml(temp_path))
      override = types.at_xpath(
        "//xmlns:Override[@PartName='/customXml/item1.xml']",
        'xmlns' => content_types_ns
      )
      expect(override['ContentType']).to eq(content_type)
    ensure
      File.delete(temp_path) if temp_path && File.exist?(temp_path)
    end
  end

  describe '#add_relationship' do
    it 'returns a unique relationship id and persists internal relationships' do
      doc = Docx::Document.open(File.join(fixtures_path, 'basic.docx'))
      existing_ids = doc.instance_variable_get(:@rels).xpath('//xmlns:Relationship/@Id').map(&:value)

      rid = doc.add_relationship(
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml',
        'customXml/item1.xml'
      )

      expect(rid).to match(/\ArId\d+\z/)
      expect(existing_ids).not_to include(rid)

      relation = doc.instance_variable_get(:@rels).at_xpath("//xmlns:Relationship[@Id='#{rid}']")
      expect(relation['Type']).to eq('http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml')
      expect(relation['Target']).to eq('customXml/item1.xml')
      expect(relation['TargetMode']).to be_nil

      temp_path = save_to_tempfile(doc)
      saved_rels = Nokogiri::XML(rels_xml(temp_path))
      saved_relation = saved_rels.at_xpath("//xmlns:Relationship[@Id='#{rid}']")
      expect(saved_relation).not_to be_nil
    ensure
      File.delete(temp_path) if temp_path && File.exist?(temp_path)
    end

    it 'persists external relationships with TargetMode' do
      doc = Docx::Document.open(File.join(fixtures_path, 'basic.docx'))

      rid = doc.add_relationship(
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        'https://example.com',
        mode: :external
      )

      relation = doc.instance_variable_get(:@rels).at_xpath("//xmlns:Relationship[@Id='#{rid}']")
      expect(relation['TargetMode']).to eq('External')

      temp_path = save_to_tempfile(doc)
      saved_rels = Nokogiri::XML(rels_xml(temp_path))
      saved_relation = saved_rels.at_xpath("//xmlns:Relationship[@Id='#{rid}']")
      expect(saved_relation['TargetMode']).to eq('External')
    ensure
      File.delete(temp_path) if temp_path && File.exist?(temp_path)
    end
  end

  describe '#ensure_default_content_type' do
    it 'registers a Default once and remains idempotent' do
      doc = Docx::Document.open(File.join(fixtures_path, 'basic.docx'))

      doc.ensure_default_content_type('emf', 'image/x-emf')
      doc.ensure_default_content_type('emf', 'image/x-emf')

      temp_path = save_to_tempfile(doc)
      types = Nokogiri::XML(content_types_xml(temp_path))
      defaults = types.xpath("//xmlns:Default[@Extension='emf']", 'xmlns' => content_types_ns)
      expect(defaults.size).to eq(1)
      expect(defaults.first['ContentType']).to eq('image/x-emf')
    ensure
      File.delete(temp_path) if temp_path && File.exist?(temp_path)
    end
  end

  describe 'styles persistence from @styles DOM' do
    it 'writes direct @styles DOM edits on save' do
      doc = Docx::Document.open(File.join(fixtures_path, 'styles.docx'))
      styles_root = doc.styles.at_xpath('//w:styles', Docx::Document::XML_NAMESPACES)
      style_node = Nokogiri::XML::Node.new('w:style', doc.styles)
      style_node['w:styleId'] = 'ZZTest'
      styles_root.add_child(style_node)

      temp_path = save_to_tempfile(doc)
      reopened = Docx::Document.open(temp_path)
      saved_style = reopened.styles.at_xpath("//w:style[@w:styleId='ZZTest']", Docx::Document::XML_NAMESPACES)
      expect(saved_style).not_to be_nil
    ensure
      File.delete(temp_path) if temp_path && File.exist?(temp_path)
    end
  end

  describe '#ensure_numbering!' do
    it 'creates numbering.xml, content type, and relationship, and is idempotent' do
      doc = Docx::Document.open(File.join(fixtures_path, 'basic.docx'))
      expect(doc.numbering).to be_nil

      first = doc.ensure_numbering!
      second = doc.ensure_numbering!
      expect(first).to equal(second)

      temp_path = save_to_tempfile(doc)

      Zip::File.open(temp_path) do |zip|
        expect(zip.find_entry('word/numbering.xml')).not_to be_nil
      end

      types = Nokogiri::XML(content_types_xml(temp_path))
      override = types.at_xpath(
        "//xmlns:Override[@PartName='/word/numbering.xml']",
        'xmlns' => content_types_ns
      )
      expect(override['ContentType']).to eq(Docx::Document::NUMBERING_CONTENT_TYPE)

      saved_rels = Nokogiri::XML(rels_xml(temp_path))
      numbering_rels = saved_rels.xpath(
        "//xmlns:Relationship[@Type='#{Docx::Document::NUMBERING_REL_TYPE}']",
        'xmlns' => rels_ns
      )
      expect(numbering_rels.size).to eq(1)
      expect(numbering_rels.first['Target']).to eq('numbering.xml')
    ensure
      File.delete(temp_path) if temp_path && File.exist?(temp_path)
    end
  end
end
