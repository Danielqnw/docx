# frozen_string_literal: true

require 'spec_helper'
require 'docx'
require 'tempfile'
require 'zip'
require 'stringio'

describe Docx::Merge::NumberingImporter do
  ns = Docx::Document::XML_NAMESPACES

  def xml_ns
    Docx::Document::XML_NAMESPACES
  end

  def build_doc_with_numbering(numbering_xml)
    base = File.join('spec/fixtures', 'basic.docx')
    buffer = Zip::OutputStream.write_buffer do |out|
      has_numbering = false
      Zip::File.open(base) do |zf|
        zf.each do |entry|
          next unless entry.file?

          if entry.name == 'word/numbering.xml'
            has_numbering = true
            out.put_next_entry(entry.name)
            out.write(numbering_xml)
          else
            out.put_next_entry(entry.name)
            out.write(zf.read(entry.name))
          end
        end
      end
      unless has_numbering
        out.put_next_entry('word/numbering.xml')
        out.write(numbering_xml)
      end
    end
    Docx::Document.open(StringIO.new(buffer.string))
  end

  def build_doc_without_numbering
    Docx::Document.open(File.join('spec/fixtures', 'basic.docx'))
  end

  def save_to_tempfile(doc)
    temp_file = Tempfile.new(['numbering_importer', '.docx'])
    temp_path = temp_file.path
    temp_file.close
    doc.save(temp_path)
    temp_path
  end

  def wrap_numbering(*children)
    w_ns = Docx::Document::XML_NAMESPACES['w']
    <<~XML
      <?xml version="1.0"?>
      <w:numbering xmlns:w="#{w_ns}">
        #{children.join("\n")}
      </w:numbering>
    XML
  end

  def abstract_num(abs_id, num_style_link: nil, lvl_text: '•')
    style_link_xml = if num_style_link
                       %(        <w:numStyleLink w:val="#{num_style_link}"/>)
                     else
                       ''
                     end
    <<~XML
      <w:abstractNum w:abstractNumId="#{abs_id}">
      #{style_link_xml}
        <w:lvl w:ilvl="0">
          <w:start w:val="1"/>
          <w:numFmt w:val="bullet"/>
          <w:lvlText w:val="#{lvl_text}"/>
        </w:lvl>
      </w:abstractNum>
    XML
  end

  def num(num_id, abstract_num_id)
    <<~XML
      <w:num w:numId="#{num_id}">
        <w:abstractNumId w:val="#{abstract_num_id}"/>
      </w:num>
    XML
  end

  def abstract_num_nodes(doc)
    doc.numbering.xpath('//w:numbering/w:abstractNum', xml_ns)
  end

  def num_nodes(doc)
    doc.numbering.xpath('//w:numbering/w:num', xml_ns)
  end

  def abstract_num_node(doc, abs_id)
    doc.numbering.at_xpath("//w:abstractNum[@w:abstractNumId='#{abs_id}']", xml_ns)
  end

  def num_node(doc, num_id)
    doc.numbering.at_xpath("//w:num[@w:numId='#{num_id}']", xml_ns)
  end

  describe '#import' do
    it 'creates numbering on target and imports with offset ids when target has none' do
      target = build_doc_without_numbering
      source = build_doc_with_numbering(wrap_numbering(abstract_num(0), num(1, 0)))

      importer = described_class.new(target, source)
      mapped_id = importer.import('1')

      expect(target.numbering).not_to be_nil
      expect(mapped_id).to eq('1')
      expect(importer.abstract_num_id_map).to eq('0' => '0')
      expect(importer.num_id_map).to eq('1' => '1')
      expect(abstract_num_node(target, '0')).not_to be_nil
      imported_num = num_node(target, '1')
      expect(imported_num).not_to be_nil
      expect(imported_num.at_xpath('./w:abstractNumId/@w:val', ns).value).to eq('0')
    end

    it 'offsets imported ids to avoid conflicts with existing target numbering' do
      target = build_doc_with_numbering(wrap_numbering(abstract_num(0, lvl_text: 'T'), num(1, 0)))
      source = build_doc_with_numbering(wrap_numbering(abstract_num(0, lvl_text: 'S'), num(1, 0)))
      target_abs_xml = abstract_num_node(target, '0').to_xml
      target_num_xml = num_node(target, '1').to_xml

      importer = described_class.new(target, source)
      mapped_id = importer.import('1')

      expect(mapped_id).to eq('2')
      expect(importer.abstract_num_id_map).to eq('0' => '1')
      expect(importer.num_id_map).to eq('1' => '2')
      expect(abstract_num_node(target, '0').to_xml).to eq(target_abs_xml)
      expect(num_node(target, '1').to_xml).to eq(target_num_xml)
      expect(abstract_num_node(target, '1').at_xpath('.//w:lvlText/@w:val', ns).value).to eq('S')
      expect(num_node(target, '2').at_xpath('./w:abstractNumId/@w:val', ns).value).to eq('1')
    end

    it 'rewrites numStyleLink values using style_id_map' do
      target = build_doc_without_numbering
      source = build_doc_with_numbering(
        wrap_numbering(abstract_num(0, num_style_link: 'SrcStyle'), num(1, 0))
      )
      style_map = { 'SrcStyle' => 'm1_SrcStyle' }

      importer = described_class.new(target, source, style_id_map: style_map)
      importer.import('1')

      imported_abs = abstract_num_node(target, '0')
      expect(imported_abs.at_xpath('./w:numStyleLink/@w:val', ns).value).to eq('m1_SrcStyle')
    end

    it 'imports shared abstractNum only once when multiple nums reference it' do
      target = build_doc_without_numbering
      source = build_doc_with_numbering(
        wrap_numbering(abstract_num(0), num(1, 0), num(2, 0))
      )
      initial_abs_count = abstract_num_nodes(source).size

      importer = described_class.new(target, source)
      importer.import('1')
      importer.import('2')

      expect(initial_abs_count).to eq(1)
      expect(abstract_num_nodes(target).size).to eq(1)
      expect(num_nodes(target).size).to eq(2)
      expect(importer.abstract_num_id_map).to eq('0' => '0')
      expect(importer.num_id_map).to eq('1' => '1', '2' => '2')
      expect(num_node(target, '1').at_xpath('./w:abstractNumId/@w:val', ns).value).to eq('0')
      expect(num_node(target, '2').at_xpath('./w:abstractNumId/@w:val', ns).value).to eq('0')
    end

    it 'keeps all abstractNum elements before num elements after import' do
      target = build_doc_without_numbering
      source = build_doc_with_numbering(
        wrap_numbering(abstract_num(0), abstract_num(1), num(1, 0), num(2, 1))
      )

      importer = described_class.new(target, source)
      importer.import_all

      children = target.numbering.at_xpath('//w:numbering', ns).children.select(&:element?)
      last_abs_index = children.rindex { |node| node.name == 'abstractNum' }
      first_num_index = children.index { |node| node.name == 'num' }
      expect(last_abs_index).to be < first_num_index
    end

    it 'returns the original id when source numbering is missing' do
      target = build_doc_without_numbering
      source = build_doc_without_numbering

      importer = described_class.new(target, source)
      mapped_id = importer.import('1')

      expect(mapped_id).to eq('1')
      expect(target.numbering).to be_nil
      expect(importer.num_id_map).to be_empty
      expect(importer.abstract_num_id_map).to be_empty
    end

    it 'returns the original id when source numId is missing' do
      target = build_doc_without_numbering
      source = build_doc_with_numbering(wrap_numbering(abstract_num(0), num(1, 0)))

      importer = described_class.new(target, source)
      mapped_id = importer.import('99')

      expect(mapped_id).to eq('99')
      expect(importer.num_id_map).to eq('99' => '99')
      expect(num_nodes(target)).to be_empty
    end

    it 'serializes imported numbering with w: namespace prefixes' do
      target = build_doc_without_numbering
      source = build_doc_with_numbering(wrap_numbering(abstract_num(0), num(1, 0)))

      importer = described_class.new(target, source)
      importer.import('1')

      serialized = target.numbering.serialize(save_with: 0)
      expect(serialized).to include('w:abstractNum')
      expect(serialized).to include('w:num')
      expect(serialized.scan('xmlns:w=').size).to eq(1)
    end

    it 'persists imported numbering after save and reopen' do
      target = build_doc_without_numbering
      source = build_doc_with_numbering(wrap_numbering(abstract_num(0), num(1, 0)))

      importer = described_class.new(target, source)
      importer.import('1')

      temp_path = save_to_tempfile(target)
      reopened = Docx::Document.open(temp_path)
      expect(num_node(reopened, '1')).not_to be_nil
      expect(abstract_num_node(reopened, '0')).not_to be_nil
    ensure
      File.delete(temp_path) if defined?(temp_path) && temp_path && File.exist?(temp_path)
    end
  end

  describe '#import_all' do
    it 'imports every num definition from source' do
      target = build_doc_without_numbering
      source = build_doc_with_numbering(
        wrap_numbering(abstract_num(0), abstract_num(1), num(1, 0), num(2, 1))
      )

      importer = described_class.new(target, source)
      importer.import_all

      expect(importer.num_id_map.keys.sort).to eq(%w[1 2])
      expect(num_nodes(target).size).to eq(2)
      expect(abstract_num_nodes(target).size).to eq(2)
    end
  end
end
