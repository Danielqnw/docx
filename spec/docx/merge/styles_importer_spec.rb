# frozen_string_literal: true

require 'spec_helper'
require 'docx'
require 'tempfile'
require 'zip'
require 'stringio'

describe Docx::Merge::StylesImporter do
  ns = Docx::Document::XML_NAMESPACES

  def build_doc_with_styles(styles_xml)
    base = File.join('spec/fixtures', 'basic.docx')
    buffer = Zip::OutputStream.write_buffer do |out|
      Zip::File.open(base) do |zf|
        zf.each do |entry|
          next unless entry.file?

          out.put_next_entry(entry.name)
          out.write(entry.name == 'word/styles.xml' ? styles_xml : zf.read(entry.name))
        end
      end
    end
    Docx::Document.open(StringIO.new(buffer.string))
  end

  def save_to_tempfile(doc)
    temp_file = Tempfile.new(['styles_importer', '.docx'])
    temp_path = temp_file.path
    temp_file.close
    doc.save(temp_path)
    temp_path
  end

  def style_nodes(doc)
    doc.styles.xpath('//w:styles/w:style', Docx::Document::XML_NAMESPACES)
  end

  def style_node(doc, style_id)
    doc.styles.at_xpath("//w:style[@w:styleId='#{style_id}']", Docx::Document::XML_NAMESPACES)
  end

  def doc_defaults_xml
    <<~XML
      <w:docDefaults>
        <w:rPrDefault>
          <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
          </w:rPr>
        </w:rPrDefault>
        <w:pPrDefault/>
      </w:docDefaults>
    XML
  end

  def tbl_style_with_border(style_id, border_val)
    <<~XML
      <w:style w:type="table" w:styleId="#{style_id}">
        <w:name w:val="#{style_id}"/>
        <w:tblPr>
          <w:tblBorders>
            <w:top w:val="#{border_val}" w:sz="4" w:space="0" w:color="auto"/>
          </w:tblBorders>
        </w:tblPr>
      </w:style>
    XML
  end

  def wrap_styles(*children)
    <<~XML
      <?xml version="1.0"?>
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        #{doc_defaults_xml}
        #{children.join("\n")}
      </w:styles>
    XML
  end

  describe '#import' do
    it 'renames conflicting table styles and preserves the target definition' do
      target = build_doc_with_styles(wrap_styles(tbl_style_with_border('TblX', 'single')))
      source = build_doc_with_styles(wrap_styles(tbl_style_with_border('TblX', 'none')))

      importer = described_class.new(target, source)
      mapped_id = importer.import('TblX')

      expect(mapped_id).to eq('m1_TblX')
      expect(style_node(target, 'TblX').at_xpath('.//w:top/@w:val', ns).value).to eq('single')
      imported = style_node(target, 'm1_TblX')
      expect(imported).not_to be_nil
      expect(imported.at_xpath('.//w:top/@w:val', ns).value).to eq('none')
      expect(importer.style_id_map).to eq('TblX' => 'm1_TblX')
    end

    it 'imports missing styles with the original styleId' do
      only_in_source = <<~XML
        <w:style w:type="paragraph" w:styleId="OnlyInSource">
          <w:name w:val="Only In Source"/>
          <w:pPr><w:spacing w:after="120"/></w:pPr>
        </w:style>
      XML
      target = build_doc_with_styles(wrap_styles)
      source = build_doc_with_styles(wrap_styles(only_in_source))

      importer = described_class.new(target, source)
      mapped_id = importer.import('OnlyInSource')

      expect(mapped_id).to eq('OnlyInSource')
      expect(style_node(target, 'OnlyInSource')).not_to be_nil
      expect(importer.style_id_map).to eq('OnlyInSource' => 'OnlyInSource')
    end

    it 'reuses equivalent styles without appending duplicates' do
      shared_style = <<~XML
        <w:style w:type="paragraph" w:styleId="SharedStyle">
          <w:name w:val="Shared"/>
          <w:pPr><w:spacing w:after="200"/></w:pPr>
        </w:style>
      XML
      target = build_doc_with_styles(wrap_styles(shared_style))
      source = build_doc_with_styles(wrap_styles(shared_style))
      initial_count = style_nodes(target).size

      importer = described_class.new(target, source)
      mapped_id = importer.import('SharedStyle')

      expect(mapped_id).to eq('SharedStyle')
      expect(style_nodes(target).size).to eq(initial_count)
      expect(importer.style_id_map).to eq('SharedStyle' => 'SharedStyle')
    end

    it 'imports dependency closures and rewrites basedOn when parent conflicts' do
      target_parent = <<~XML
        <w:style w:type="paragraph" w:styleId="Parent">
          <w:name w:val="Parent Target"/>
          <w:pPr><w:spacing w:after="100"/></w:pPr>
        </w:style>
      XML
      source_parent = <<~XML
        <w:style w:type="paragraph" w:styleId="Parent">
          <w:name w:val="Parent Source"/>
          <w:pPr><w:spacing w:after="300"/></w:pPr>
        </w:style>
      XML
      source_child = <<~XML
        <w:style w:type="paragraph" w:styleId="Child">
          <w:name w:val="Child"/>
          <w:basedOn w:val="Parent"/>
          <w:pPr><w:spacing w:after="400"/></w:pPr>
        </w:style>
      XML

      target = build_doc_with_styles(wrap_styles(target_parent))
      source = build_doc_with_styles(wrap_styles(source_parent, source_child))

      importer = described_class.new(target, source)
      child_id = importer.import('Child')

      expect(child_id).to eq('Child')
      expect(importer.style_id_map['Parent']).to eq('m1_Parent')
      child_node = style_node(target, 'Child')
      expect(child_node.at_xpath('w:basedOn/@w:val', ns).value).to eq('m1_Parent')
      expect(style_node(target, 'Parent').at_xpath('w:pPr/w:spacing/@w:after', ns).value).to eq('100')
      expect(style_node(target, 'm1_Parent').at_xpath('w:pPr/w:spacing/@w:after', ns).value).to eq('300')
    end

    it 'persists imported styles after save and reopen' do
      only_in_source = <<~XML
        <w:style w:type="paragraph" w:styleId="PersistMe">
          <w:name w:val="Persist"/>
        </w:style>
      XML
      target = build_doc_with_styles(wrap_styles)
      source = build_doc_with_styles(wrap_styles(only_in_source))
      importer = described_class.new(target, source)
      importer.import('PersistMe')

      temp_path = save_to_tempfile(target)
      reopened = Docx::Document.open(temp_path)
      expect(style_node(reopened, 'PersistMe')).not_to be_nil
    ensure
      File.delete(temp_path) if defined?(temp_path) && temp_path && File.exist?(temp_path)
    end

    it 'returns the original id when the source style is missing' do
      target = build_doc_with_styles(wrap_styles)
      source = build_doc_with_styles(wrap_styles)
      importer = described_class.new(target, source)
      initial_count = style_nodes(target).size

      mapped_id = importer.import('MissingStyle')

      expect(mapped_id).to eq('MissingStyle')
      expect(style_nodes(target).size).to eq(initial_count)
      expect(importer.style_id_map).to eq('MissingStyle' => 'MissingStyle')
    end

    it 'serializes imported styles with w: namespace prefixes' do
      only_in_source = <<~XML
        <w:style w:type="paragraph" w:styleId="NamespaceCheck">
          <w:name w:val="Namespace"/>
        </w:style>
      XML
      target = build_doc_with_styles(wrap_styles)
      source = build_doc_with_styles(wrap_styles(only_in_source))
      importer = described_class.new(target, source)
      importer.import('NamespaceCheck')

      serialized = target.styles.serialize(save_with: 0)
      expect(serialized).to include('w:style')
      expect(serialized).to include('w:styleId="NamespaceCheck"')
      expect(serialized.scan('xmlns:w=').size).to eq(1)
    end
  end
end
