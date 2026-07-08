# frozen_string_literal: true

require 'spec_helper'
require 'docx'
require 'securerandom'
require 'tmpdir'
require 'zip'
require 'stringio'

describe Docx::Merge::Importer do
  def xml_ns
    Docx::Document::XML_NAMESPACES
  end

  def build_doc(document_xml:, styles_xml:, numbering_xml: nil)
    base = File.join('spec/fixtures', 'basic.docx')
    buffer = Zip::OutputStream.write_buffer do |out|
      has_numbering = false
      Zip::File.open(base) do |zf|
        zf.each do |entry|
          next unless entry.file?

          content = case entry.name
                    when 'word/document.xml' then document_xml
                    when 'word/styles.xml' then styles_xml
                    when 'word/numbering.xml'
                      has_numbering = true
                      numbering_xml || zf.read(entry.name)
                    else
                      zf.read(entry.name)
                    end
          out.put_next_entry(entry.name)
          out.write(content)
        end
      end
      if numbering_xml && !has_numbering
        out.put_next_entry('word/numbering.xml')
        out.write(numbering_xml)
      end
    end
    Docx::Document.open(StringIO.new(buffer.string))
  end

  def save_to_tempfile(doc)
    path = File.join(Dir.tmpdir, "docx_merge_#{SecureRandom.hex(8)}.docx")
    doc.save(path)
    path
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

  def wrap_numbering(*children)
    w_ns = Docx::Document::XML_NAMESPACES['w']
    <<~XML
      <?xml version="1.0"?>
      <w:numbering xmlns:w="#{w_ns}">
        #{children.join("\n")}
      </w:numbering>
    XML
  end

  def abstract_num(abs_id, lvl_text: '•')
    <<~XML
      <w:abstractNum w:abstractNumId="#{abs_id}">
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

  def source_document_xml(tbl_style: 'TblX', num_id: nil)
    num_pr_xml = if num_id
                   <<~XML
                     <w:pPr>
                       <w:numPr>
                         <w:numId w:val="#{num_id}"/>
                       </w:numPr>
                     </w:pPr>
                   XML
                 else
                   ''
                 end
    <<~XML
      <?xml version="1.0"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:tbl>
            <w:tblPr>
              <w:tblStyle w:val="#{tbl_style}"/>
            </w:tblPr>
            <w:tr>
              <w:tc>
                <w:p>
                  #{num_pr_xml}
                  <w:r><w:t>cell</w:t></w:r>
                </w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
          <w:p/>
        </w:body>
      </w:document>
    XML
  end

  def target_document_xml
    <<~XML
      <?xml version="1.0"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p><w:r><w:t>anchor</w:t></w:r></w:p>
          <w:sectPr/>
        </w:body>
      </w:document>
    XML
  end

  def style_node(doc, style_id)
    doc.styles.at_xpath("//w:style[@w:styleId='#{style_id}']", xml_ns)
  end

  def style_nodes(doc)
    doc.styles.xpath('//w:styles/w:style', xml_ns)
  end

  def num_node(doc, num_id)
    doc.numbering&.at_xpath("//w:num[@w:numId='#{num_id}']", xml_ns)
  end

  def source_tbl_node(source)
    source.doc.at_xpath('//w:body/w:tbl', xml_ns)
  end

  describe '#import' do
    it 'isolates conflicting table style ids and preserves the target definition' do
      target = build_doc(
        document_xml: target_document_xml,
        styles_xml: wrap_styles(tbl_style_with_border('TblX', 'single'))
      )
      source = build_doc(
        document_xml: source_document_xml,
        styles_xml: wrap_styles(tbl_style_with_border('TblX', 'none'))
      )

      importer = described_class.new(target, source)
      imported = importer.import(source_tbl_node(source))

      mapped_id = importer.style_id_map['TblX']
      expect(mapped_id).to eq('m1_TblX')
      expect(imported.at_xpath('.//w:tblStyle/@w:val', xml_ns).value).to eq('m1_TblX')
      expect(style_node(target, 'TblX').at_xpath('.//w:top/@w:val', xml_ns).value).to eq('single')
      expect(style_node(target, 'm1_TblX').at_xpath('.//w:top/@w:val', xml_ns).value).to eq('none')

      anchor = target.doc.at_xpath('//w:body/w:p', xml_ns)
      anchor.add_previous_sibling(imported)
      expect(target.doc.at_xpath('//w:body/w:tbl', xml_ns)).not_to be_nil
      expect(style_node(target, mapped_id)).not_to be_nil
    end

    it 'imports missing styles with the original styleId' do
      only_in_source = <<~XML
        <w:style w:type="paragraph" w:styleId="OnlyInSource">
          <w:name w:val="Only In Source"/>
          <w:pPr><w:spacing w:after="120"/></w:pPr>
        </w:style>
      XML
      target = build_doc(document_xml: target_document_xml, styles_xml: wrap_styles)
      source = build_doc(
        document_xml: source_document_xml(tbl_style: 'OnlyInSource'),
        styles_xml: wrap_styles(only_in_source)
      )

      importer = described_class.new(target, source)
      imported = importer.import(source_tbl_node(source))

      expect(imported.at_xpath('.//w:tblStyle/@w:val', xml_ns).value).to eq('OnlyInSource')
      expect(style_node(target, 'OnlyInSource')).not_to be_nil
      expect(importer.style_id_map).to eq('OnlyInSource' => 'OnlyInSource')
    end

    it 'reuses equivalent styles without appending duplicates' do
      shared_style = tbl_style_with_border('TblX', 'single')
      target = build_doc(
        document_xml: target_document_xml,
        styles_xml: wrap_styles(shared_style)
      )
      source = build_doc(
        document_xml: source_document_xml,
        styles_xml: wrap_styles(shared_style)
      )
      initial_count = style_nodes(target).size

      importer = described_class.new(target, source)
      imported = importer.import(source_tbl_node(source))

      expect(imported.at_xpath('.//w:tblStyle/@w:val', xml_ns).value).to eq('TblX')
      expect(style_nodes(target).size).to eq(initial_count)
      expect(importer.style_id_map).to eq('TblX' => 'TblX')
    end

    it 'imports numbering definitions and rewrites numId references' do
      target = build_doc(
        document_xml: target_document_xml,
        styles_xml: wrap_styles
      )
      source = build_doc(
        document_xml: source_document_xml(num_id: '1'),
        styles_xml: wrap_styles,
        numbering_xml: wrap_numbering(abstract_num(0, lvl_text: 'S'), num(1, 0))
      )

      importer = described_class.new(target, source)
      imported = importer.import(source_tbl_node(source))

      expect(target.numbering).not_to be_nil
      expect(imported.at_xpath('.//w:numPr/w:numId/@w:val', xml_ns).value).to eq('1')
      expect(importer.num_id_map).to eq('1' => '1')
      expect(num_node(target, '1')).not_to be_nil
      expect(target.numbering.at_xpath("//w:abstractNum[@w:abstractNumId='0']", xml_ns)).not_to be_nil
    end

    it 'reuses style imports across multiple nodes referencing the same style' do
      document_xml = <<~XML
        <?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:tbl>
              <w:tblPr><w:tblStyle w:val="TblX"/></w:tblPr>
              <w:tr><w:tc><w:p><w:r><w:t>one</w:t></w:r></w:p></w:tc></w:tr>
            </w:tbl>
            <w:tbl>
              <w:tblPr><w:tblStyle w:val="TblX"/></w:tblPr>
              <w:tr><w:tc><w:p><w:r><w:t>two</w:t></w:r></w:p></w:tc></w:tr>
            </w:tbl>
          </w:body>
        </w:document>
      XML
      target = build_doc(
        document_xml: target_document_xml,
        styles_xml: wrap_styles(tbl_style_with_border('TblX', 'single'))
      )
      source = build_doc(
        document_xml: document_xml,
        styles_xml: wrap_styles(tbl_style_with_border('TblX', 'none'))
      )
      tbl_nodes = source.doc.xpath('//w:body/w:tbl', xml_ns)

      importer = described_class.new(target, source)
      first = importer.import(tbl_nodes[0])
      count_after_first = style_nodes(target).size
      second = importer.import(tbl_nodes[1])

      expect(style_nodes(target).size).to eq(count_after_first)
      expect(first.at_xpath('.//w:tblStyle/@w:val', xml_ns).value).to eq('m1_TblX')
      expect(second.at_xpath('.//w:tblStyle/@w:val', xml_ns).value).to eq('m1_TblX')
      expect(style_nodes(target).count { |s| s['w:styleId'] == 'm1_TblX' }).to eq(1)
    end

    it 'does not modify the original source node' do
      target = build_doc(
        document_xml: target_document_xml,
        styles_xml: wrap_styles(tbl_style_with_border('TblX', 'single'))
      )
      source = build_doc(
        document_xml: source_document_xml,
        styles_xml: wrap_styles(tbl_style_with_border('TblX', 'none'))
      )
      source_tbl = source_tbl_node(source)
      original_val = source_tbl.at_xpath('.//w:tblStyle/@w:val', xml_ns).value

      described_class.new(target, source).import(source_tbl)

      expect(source_tbl.at_xpath('.//w:tblStyle/@w:val', xml_ns).value).to eq(original_val)
    end

    it 'produces valid style and numbering references after save and reopen' do
      target = build_doc(
        document_xml: target_document_xml,
        styles_xml: wrap_styles(tbl_style_with_border('TblX', 'single'))
      )
      source = build_doc(
        document_xml: source_document_xml(num_id: '1'),
        styles_xml: wrap_styles(tbl_style_with_border('TblX', 'none')),
        numbering_xml: wrap_numbering(abstract_num(0), num(1, 0))
      )

      importer = described_class.new(target, source)
      imported = importer.import(source_tbl_node(source))
      anchor = target.doc.at_xpath('//w:body/w:p', xml_ns)
      anchor.add_previous_sibling(imported)

      temp_path = save_to_tempfile(target)
      reopened = Docx::Document.open(temp_path)

      reopened.doc.xpath('//w:body//w:pStyle | //w:body//w:rStyle | //w:body//w:tblStyle', xml_ns).each do |style_ref|
        style_id = style_ref['w:val']
        expect(style_node(reopened, style_id)).not_to be_nil, "missing style #{style_id}"
      end

      reopened.doc.xpath('//w:body//w:numPr/w:numId', xml_ns).each do |num_ref|
        num_id = num_ref['w:val']
        expect(num_node(reopened, num_id)).not_to be_nil, "missing num #{num_id}"
      end

      expect(reopened.doc.at_xpath('//w:body/w:tbl', xml_ns)).not_to be_nil
      serialized = reopened.doc.serialize(save_with: 0)
      expect(serialized.scan('xmlns:w=').size).to eq(1)
    ensure
      File.delete(temp_path) if defined?(temp_path) && temp_path && File.exist?(temp_path)
    end
  end
end

describe 'Docx::Document cross-document import API' do
  def xml_ns
    Docx::Document::XML_NAMESPACES
  end

  def build_doc(document_xml:, styles_xml:, numbering_xml: nil)
    base = File.join('spec/fixtures', 'basic.docx')
    buffer = Zip::OutputStream.write_buffer do |out|
      has_numbering = false
      Zip::File.open(base) do |zf|
        zf.each do |entry|
          next unless entry.file?

          content = case entry.name
                    when 'word/document.xml' then document_xml
                    when 'word/styles.xml' then styles_xml
                    when 'word/numbering.xml'
                      has_numbering = true
                      numbering_xml || zf.read(entry.name)
                    else
                      zf.read(entry.name)
                    end
          out.put_next_entry(entry.name)
          out.write(content)
        end
      end
      if numbering_xml && !has_numbering
        out.put_next_entry('word/numbering.xml')
        out.write(numbering_xml)
      end
    end
    Docx::Document.open(StringIO.new(buffer.string))
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

  def source_document_xml
    <<~XML
      <?xml version="1.0"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:tbl>
            <w:tblPr><w:tblStyle w:val="TblX"/></w:tblPr>
            <w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr>
          </w:tbl>
        </w:body>
      </w:document>
    XML
  end

  def target_document_xml
    <<~XML
      <?xml version="1.0"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p><w:r><w:t>anchor</w:t></w:r></w:p>
          <w:sectPr/>
        </w:body>
      </w:document>
    XML
  end

  def style_nodes(doc)
    doc.styles.xpath('//w:styles/w:style', xml_ns)
  end

  describe '#import_before' do
    it 'inserts an isolated node before the anchor and reuses importer per source' do
      target = build_doc(
        document_xml: target_document_xml,
        styles_xml: wrap_styles(tbl_style_with_border('TblX', 'single'))
      )
      source = build_doc(
        document_xml: source_document_xml,
        styles_xml: wrap_styles(tbl_style_with_border('TblX', 'none'))
      )
      source_tbl = source.doc.at_xpath('//w:body/w:tbl', xml_ns)
      anchor = target.doc.at_xpath('//w:body/w:p', xml_ns)

      imported = target.import_before(source, source_tbl, anchor)

      body_children = target.doc.at_xpath('//w:body', xml_ns).element_children.reject { |n| n.name == 'sectPr' }
      expect(body_children.map(&:name)).to eq(%w[tbl p])
      expect(imported.at_xpath('.//w:tblStyle/@w:val', xml_ns).value).to eq('m1_TblX')
      expect(style_nodes(target).count { |s| s['w:styleId'] == 'm1_TblX' }).to eq(1)

      second_tbl = source_tbl.dup(1)
      count_after_first = style_nodes(target).size
      target.import_node(source, second_tbl)
      expect(style_nodes(target).size).to eq(count_after_first)

      cached = target.instance_variable_get(:@merge_importers)[source]
      target.import_node(source, second_tbl)
      expect(target.instance_variable_get(:@merge_importers)[source]).to equal(cached)
    end
  end

  describe '#import_after' do
    it 'inserts an isolated node after the anchor' do
      target = build_doc(
        document_xml: target_document_xml,
        styles_xml: wrap_styles(tbl_style_with_border('TblX', 'single'))
      )
      source = build_doc(
        document_xml: source_document_xml,
        styles_xml: wrap_styles(tbl_style_with_border('TblX', 'none'))
      )
      source_tbl = source.doc.at_xpath('//w:body/w:tbl', xml_ns)
      anchor = target.doc.at_xpath('//w:body/w:p', xml_ns)

      imported = target.import_after(source, source_tbl, anchor)

      body_children = target.doc.at_xpath('//w:body', xml_ns).element_children.reject { |n| n.name == 'sectPr' }
      expect(body_children.map(&:name)).to eq(%w[p tbl])
      expect(imported.at_xpath('.//w:tblStyle/@w:val', xml_ns).value).to eq('m1_TblX')
    end
  end

  describe '#import_node fallback' do
    it 'returns a dup without raising when fallback is true' do
      target = build_doc(document_xml: target_document_xml, styles_xml: wrap_styles)
      source = build_doc(document_xml: source_document_xml, styles_xml: wrap_styles)
      source_tbl = source.doc.at_xpath('//w:body/w:tbl', xml_ns)

      allow_any_instance_of(Docx::Merge::Importer).to receive(:import).and_raise(StandardError, 'boom')

      expect do
        result = target.import_node(source, source_tbl, fallback: true)
        expect(result.at_xpath('.//w:tblStyle/@w:val', xml_ns).value).to eq('TblX')
      end.not_to raise_error
    end

    it 're-raises when fallback is false' do
      target = build_doc(document_xml: target_document_xml, styles_xml: wrap_styles)
      source = build_doc(document_xml: source_document_xml, styles_xml: wrap_styles)
      source_tbl = source.doc.at_xpath('//w:body/w:tbl', xml_ns)

      allow_any_instance_of(Docx::Merge::Importer).to receive(:import).and_raise(StandardError, 'boom')

      expect do
        target.import_node(source, source_tbl, fallback: false)
      end.to raise_error(StandardError, 'boom')
    end
  end
end
