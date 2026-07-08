# frozen_string_literal: true

require 'spec_helper'
require 'docx'
require 'securerandom'
require 'tmpdir'
require 'zip'
require 'stringio'

describe Docx::Merge::Importer do
  FIXTURES_PATH = 'spec/fixtures'
  RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
  CONTENT_TYPES_NS = Docx::Document::CONTENT_TYPES_NS
  IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
  SOURCE_IMAGE_RID = 'rId100'
  SOURCE_MEDIA_TARGET = 'media/image1.png'
  SOURCE_MEDIA_PATH = "word/#{SOURCE_MEDIA_TARGET}".freeze

  def xml_ns
    Docx::Document::XML_NAMESPACES
  end

  def drawing_ns
    {
      'a' => 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
  end

  def append_image_relationship(rels_xml, rid:, target:)
    insertion = %(<Relationship Id="#{rid}" Type="#{IMAGE_REL_TYPE}" Target="#{target}"/>)
    rels_xml.sub('</Relationships>', "#{insertion}</Relationships>")
  end

  def source_document_xml_with_image(embed_rid: SOURCE_IMAGE_RID)
    <<~XML
      <?xml version="1.0"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <w:body>
          <w:p><w:r><w:drawing><wp:inline>
            <wp:extent cx="100" cy="100"/>
            <wp:docPr id="1" name="Pic"/>
            <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic>
                <pic:nvPicPr><pic:cNvPr id="1" name="Pic"/><pic:cNvPicPr/></pic:nvPicPr>
                <pic:blipFill><a:blip r:embed="#{embed_rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
              </pic:pic>
            </a:graphicData></a:graphic>
          </wp:inline></w:drawing></w:r></w:p>
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

  def build_source_with_image(document_xml: source_document_xml_with_image, image_bytes: replacement_bytes)
    base = File.join(FIXTURES_PATH, 'basic.docx')
    buffer = Zip::OutputStream.write_buffer do |out|
      Zip::File.open(base) do |zf|
        zf.each do |entry|
          next unless entry.file?

          content = case entry.name
                    when 'word/document.xml' then document_xml
                    when 'word/_rels/document.xml.rels'
                      append_image_relationship(zf.read(entry.name), rid: SOURCE_IMAGE_RID, target: SOURCE_MEDIA_TARGET)
                    else
                      zf.read(entry.name)
                    end
          out.put_next_entry(entry.name)
          out.write(content)
        end

        out.put_next_entry(SOURCE_MEDIA_PATH)
        out.write(image_bytes)
      end
    end
    Docx::Document.open(StringIO.new(buffer.string))
  end

  def build_target(document_xml: target_document_xml)
    base = File.join(FIXTURES_PATH, 'basic.docx')
    buffer = Zip::OutputStream.write_buffer do |out|
      Zip::File.open(base) do |zf|
        zf.each do |entry|
          next unless entry.file?

          content = case entry.name
                    when 'word/document.xml' then document_xml
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

  def source_image_paragraph(source)
    source.doc.at_xpath('//w:body/w:p[.//w:drawing]', xml_ns)
  end

  def blip_embed(node)
    node.at_xpath('.//a:blip/@r:embed', drawing_ns)&.value
  end

  let(:replacement_bytes) { File.binread(File.join(FIXTURES_PATH, 'replacement.png')) }

  describe '#import with embedded images' do
    it 'imports media relationships and rewrites embed references' do
      target = build_target
      source = build_source_with_image
      source_p = source_image_paragraph(source)
      existing_rids = target.instance_variable_get(:@rels).xpath('//xmlns:Relationship/@Id', 'xmlns' => RELS_NS).map(&:value)

      importer = described_class.new(target, source)
      imported = importer.import(source_p)

      new_rid = blip_embed(imported)
      expect(new_rid).to match(/\ArId\d+\z/)
      expect(new_rid).not_to eq(SOURCE_IMAGE_RID)
      expect(existing_rids).not_to include(new_rid)
      expect(importer.rid_map[SOURCE_IMAGE_RID]).to eq(new_rid)

      replace = target.instance_variable_get(:@replace)
      new_media_entry = replace.keys.find { |path| path.match?(%r{\Aword/media/image\d+\.png\z}) }
      expect(new_media_entry).not_to be_nil
      expect(replace[new_media_entry]).to eq(replacement_bytes)

      relation = target.instance_variable_get(:@rels).at_xpath("//xmlns:Relationship[@Id='#{new_rid}']", 'xmlns' => RELS_NS)
      expect(relation['Type']).to eq(IMAGE_REL_TYPE)
      expect(relation['Target']).to eq(new_media_entry.sub(%r{\Aword/}, ''))

      content_types = Nokogiri::XML(replace[Docx::Document::CONTENT_TYPES_PATH])
      png_default = content_types.at_xpath("//xmlns:Default[@Extension='png']", 'xmlns' => CONTENT_TYPES_NS)
      expect(png_default['ContentType']).to eq('image/png')
    end

    it 'produces valid media and relationship references after save and reopen' do
      target = build_target
      source = build_source_with_image
      source_p = source_image_paragraph(source)

      importer = described_class.new(target, source)
      imported = importer.import(source_p)
      new_rid = blip_embed(imported)
      new_media_entry = target.instance_variable_get(:@replace).keys.find { |path| path.match?(%r{\Aword/media/image\d+\.png\z}) }

      anchor = target.doc.at_xpath('//w:body/w:p', xml_ns)
      anchor.add_previous_sibling(imported)

      temp_path = save_to_tempfile(target)
      reopened = Docx::Document.open(temp_path)

      Zip::File.open(temp_path) do |zip|
        expect(zip.find_entry(new_media_entry)).not_to be_nil
        expect(zip.read(new_media_entry)).to eq(replacement_bytes)
      end

      reopened_embed = reopened.doc.at_xpath('//a:blip/@r:embed', drawing_ns).value
      expect(reopened_embed).to eq(new_rid)

      saved_rels = Nokogiri::XML(Zip::File.open(temp_path) { |zip| zip.read('word/_rels/document.xml.rels') })
      image_rel = saved_rels.at_xpath("//xmlns:Relationship[@Id='#{reopened_embed}']", 'xmlns' => RELS_NS)
      expect(image_rel).not_to be_nil
      expect(image_rel['Type']).to eq(IMAGE_REL_TYPE)

      resolved_path = reopened.images[reopened_embed]
      expect(resolved_path).to eq(new_media_entry)

      serialized = reopened.doc.serialize(save_with: 0)
      expect(serialized).to include('xmlns:r=')
      expect(serialized).to include('xmlns:a=')
      expect(serialized.scan('xmlns:w=').size).to eq(1)
    ensure
      File.delete(temp_path) if defined?(temp_path) && temp_path && File.exist?(temp_path)
    end

    it 'deduplicates media when importing multiple nodes referencing the same image' do
      document_xml = <<~XML
        <?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <w:body>
            <w:p><w:r><w:drawing><wp:inline>
              <wp:extent cx="100" cy="100"/>
              <wp:docPr id="1" name="Pic1"/>
              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr><pic:cNvPr id="1" name="Pic1"/><pic:cNvPicPr/></pic:nvPicPr>
                  <pic:blipFill><a:blip r:embed="#{SOURCE_IMAGE_RID}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                  <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                </pic:pic>
              </a:graphicData></a:graphic>
            </wp:inline></w:drawing></w:r></w:p>
            <w:p><w:r><w:drawing><wp:inline>
              <wp:extent cx="100" cy="100"/>
              <wp:docPr id="2" name="Pic2"/>
              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr><pic:cNvPr id="2" name="Pic2"/><pic:cNvPicPr/></pic:nvPicPr>
                  <pic:blipFill><a:blip r:embed="#{SOURCE_IMAGE_RID}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                  <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                </pic:pic>
              </a:graphicData></a:graphic>
            </wp:inline></w:drawing></w:r></w:p>
          </w:body>
        </w:document>
      XML

      target = build_target
      source = build_source_with_image(document_xml: document_xml)
      paragraphs = source.doc.xpath('//w:body/w:p[.//w:drawing]', xml_ns)

      importer = described_class.new(target, source)
      first = importer.import(paragraphs[0])
      media_count_after_first = target.instance_variable_get(:@replace).keys.grep(%r{\Aword/media/image\d+\.png\z}).size
      second = importer.import(paragraphs[1])

      expect(blip_embed(first)).to eq(blip_embed(second))
      media_entries = target.instance_variable_get(:@replace).keys.grep(%r{\Aword/media/image\d+\.png\z})
      expect(media_entries.size).to eq(media_count_after_first)
      expect(media_entries.size).to eq(1)
    end
  end

  describe 'Docx::Document cross-document image import API' do
    describe '#import_before' do
      it 'isolates embedded images with media parts and rewritten embed references' do
        target = build_target
        source = build_source_with_image
        source_p = source_image_paragraph(source)
        anchor = target.doc.at_xpath('//w:body/w:p', xml_ns)

        imported = target.import_before(source, source_p, anchor)

        new_rid = blip_embed(imported)
        expect(new_rid).not_to eq(SOURCE_IMAGE_RID)

        new_media_entry = target.instance_variable_get(:@replace).keys.find { |path| path.match?(%r{\Aword/media/image\d+\.png\z}) }
        expect(new_media_entry).not_to be_nil
        expect(target.instance_variable_get(:@replace)[new_media_entry]).to eq(replacement_bytes)

        relation = target.instance_variable_get(:@rels).at_xpath("//xmlns:Relationship[@Id='#{new_rid}']", 'xmlns' => RELS_NS)
        expect(relation['Type']).to eq(IMAGE_REL_TYPE)
        expect(relation['Target']).to eq(new_media_entry.sub(%r{\Aword/}, ''))

        temp_path = save_to_tempfile(target)
        reopened = Docx::Document.open(temp_path)

        expect(reopened.images[new_rid]).to eq(new_media_entry)
        Zip::File.open(temp_path) do |zip|
          expect(zip.read(new_media_entry)).to eq(replacement_bytes)
        end
      ensure
        File.delete(temp_path) if defined?(temp_path) && temp_path && File.exist?(temp_path)
      end
    end
  end
end
