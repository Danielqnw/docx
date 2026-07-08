# frozen_string_literal: true

require 'spec_helper'
require 'docx'
require 'securerandom'
require 'tmpdir'
require 'zip'
require 'stringio'

describe Docx::Merge::Importer, 'bookmark id isolation' do
  def xml_ns
    Docx::Document::XML_NAMESPACES
  end

  def w_ns
    xml_ns['w']
  end

  def build_doc(document_xml:, styles_xml:)
    base = File.join('spec/fixtures', 'basic.docx')
    buffer = Zip::OutputStream.write_buffer do |out|
      Zip::File.open(base) do |zf|
        zf.each do |entry|
          next unless entry.file?

          content = case entry.name
                    when 'word/document.xml' then document_xml
                    when 'word/styles.xml' then styles_xml
                    else
                      zf.read(entry.name)
                    end
          out.put_next_entry(entry.name)
          out.write(content)
        end
      end
    end
    Docx::Document.open(StringIO.new(buffer.string))
  end

  def save_to_tempfile(doc)
    path = File.join(Dir.tmpdir, "docx_merge_#{SecureRandom.hex(8)}.docx")
    doc.save(path)
    path
  end

  def wrap_styles
    <<~XML
      <?xml version="1.0"?>
      <w:styles xmlns:w="#{w_ns}">
        <w:docDefaults>
          <w:rPrDefault>
            <w:rPr>
              <w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            </w:rPr>
          </w:rPrDefault>
          <w:pPrDefault/>
        </w:docDefaults>
      </w:styles>
    XML
  end

  def target_document_xml(bookmark_id: '0', bookmark_name: 'target_bm')
    <<~XML
      <?xml version="1.0"?>
      <w:document xmlns:w="#{w_ns}">
        <w:body>
          <w:p>
            <w:bookmarkStart w:id="#{bookmark_id}" w:name="#{bookmark_name}"/>
            <w:r><w:t>target text</w:t></w:r>
            <w:bookmarkEnd w:id="#{bookmark_id}"/>
          </w:p>
          <w:sectPr/>
        </w:body>
      </w:document>
    XML
  end

  def source_document_xml(bookmark_id: '0', bookmark_name: 'ref', text: 'imported')
    <<~XML
      <?xml version="1.0"?>
      <w:document xmlns:w="#{w_ns}">
        <w:body>
          <w:p>
            <w:bookmarkStart w:id="#{bookmark_id}" w:name="#{bookmark_name}"/>
            <w:r><w:t>#{text}</w:t></w:r>
            <w:bookmarkEnd w:id="#{bookmark_id}"/>
          </w:p>
        </w:body>
      </w:document>
    XML
  end

  def source_document_xml_two_bookmarks
    <<~XML
      <?xml version="1.0"?>
      <w:document xmlns:w="#{w_ns}">
        <w:body>
          <w:p>
            <w:bookmarkStart w:id="0" w:name="first"/>
            <w:r><w:t>first text</w:t></w:r>
            <w:bookmarkEnd w:id="0"/>
          </w:p>
          <w:p>
            <w:bookmarkStart w:id="1" w:name="second"/>
            <w:r><w:t>second text</w:t></w:r>
            <w:bookmarkEnd w:id="1"/>
          </w:p>
        </w:body>
      </w:document>
    XML
  end

  def bookmark_ids(doc)
    starts = doc.doc.xpath('//w:bookmarkStart/@w:id', xml_ns).map(&:value)
    ends = doc.doc.xpath('//w:bookmarkEnd/@w:id', xml_ns).map(&:value)
    { starts: starts, ends: ends, all: (starts + ends).uniq }
  end

  def max_target_bookmark_id(doc)
    ids = doc.doc.xpath(
      '//w:bookmarkStart/@w:id | //w:bookmarkEnd/@w:id',
      xml_ns
    ).map(&:value).select { |val| val.match?(/\A\d+\z/) }.map(&:to_i)
    ids.empty? ? nil : ids.max
  end

  def paired_bookmark_ids(node)
    start_id = node.at_xpath('.//w:bookmarkStart', xml_ns)&.[]('w:id')
    end_id = node.at_xpath('.//w:bookmarkEnd', xml_ns)&.[]('w:id')
    [start_id, end_id]
  end

  describe '#import' do
    it 'offsets imported bookmark ids to avoid conflicts with target bookmarks' do
      target = build_doc(
        document_xml: target_document_xml(bookmark_id: '0'),
        styles_xml: wrap_styles
      )
      source = build_doc(
        document_xml: source_document_xml(bookmark_id: '0', bookmark_name: 'ref'),
        styles_xml: wrap_styles
      )
      source_p = source.doc.at_xpath('//w:body/w:p', xml_ns)

      importer = described_class.new(target, source)
      expect(importer.bookmark_id_offset).to eq(1)

      imported = importer.import(source_p)
      start_id, end_id = paired_bookmark_ids(imported)

      expect(start_id).to eq('1')
      expect(end_id).to eq('1')
      expect(start_id).to eq(end_id)
      expect(start_id).not_to eq('0')
      expect(bookmark_ids(target)[:all]).to include('0')
    end

    it 'exposes bookmark_id_offset as target max bookmark id plus one' do
      target = build_doc(
        document_xml: target_document_xml(bookmark_id: '5'),
        styles_xml: wrap_styles
      )
      source = build_doc(
        document_xml: source_document_xml,
        styles_xml: wrap_styles
      )

      importer = described_class.new(target, source)

      expect(importer.bookmark_id_offset).to eq(max_target_bookmark_id(target) + 1)
    end

    it 'preserves non-conflicting bookmark ids after save and reopen' do
      target = build_doc(
        document_xml: target_document_xml(bookmark_id: '0', bookmark_name: 'target_bm'),
        styles_xml: wrap_styles
      )
      source = build_doc(
        document_xml: source_document_xml(bookmark_id: '0', bookmark_name: 'ref'),
        styles_xml: wrap_styles
      )
      source_p = source.doc.at_xpath('//w:body/w:p', xml_ns)
      importer = described_class.new(target, source)
      imported = importer.import(source_p)

      anchor = target.doc.at_xpath('//w:body/w:p', xml_ns)
      anchor.add_previous_sibling(imported)

      temp_path = save_to_tempfile(target)
      reopened = Docx::Document.open(temp_path)

      target_ids = bookmark_ids(reopened)
      expect(target_ids[:starts]).to include('0', '1')
      expect(target_ids[:ends]).to include('0', '1')
      expect(target_ids[:starts].sort).to eq(target_ids[:ends].sort)

      reopened.doc.xpath('//w:bookmarkStart', xml_ns).each do |start_node|
        id = start_node['w:id']
        end_node = reopened.doc.at_xpath("//w:bookmarkEnd[@w:id='#{id}']", xml_ns)
        expect(end_node).not_to be_nil, "missing bookmarkEnd for id #{id}"
      end

      expect(reopened.bookmarks['target_bm']).not_to be_nil
      expect(reopened.bookmarks['ref']).not_to be_nil

      serialized = reopened.doc.serialize(save_with: 0)
      expect(serialized.scan('xmlns:w=').size).to eq(1)
    ensure
      File.delete(temp_path) if defined?(temp_path) && temp_path && File.exist?(temp_path)
    end

    it 'keeps bookmark ids globally unique across multiple imports from the same source' do
      target = build_doc(
        document_xml: target_document_xml(bookmark_id: '0'),
        styles_xml: wrap_styles
      )
      source = build_doc(
        document_xml: source_document_xml_two_bookmarks,
        styles_xml: wrap_styles
      )
      source_paragraphs = source.doc.xpath('//w:body/w:p', xml_ns)

      importer = described_class.new(target, source)
      offset = importer.bookmark_id_offset
      first = importer.import(source_paragraphs[0])
      second = importer.import(source_paragraphs[1])

      first_start, first_end = paired_bookmark_ids(first)
      second_start, second_end = paired_bookmark_ids(second)

      expect(first_start).to eq(first_end)
      expect(second_start).to eq(second_end)
      expect(first_start).to eq((0 + offset).to_s)
      expect(second_start).to eq((1 + offset).to_s)
      expect(first_start).not_to eq(second_start)

      anchor = target.doc.at_xpath('//w:body/w:p', xml_ns)
      anchor.add_previous_sibling(first)
      anchor.add_previous_sibling(second)

      all_ids = bookmark_ids(target)[:all]
      expect(all_ids.uniq.size).to eq(all_ids.size)
      expect(all_ids).to include('0', first_start, second_start)
    end
  end
end
