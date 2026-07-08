require 'docx/containers/text_run'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class Paragraph
        include Container
        include Elements::Element

        def self.tag
          'p'
        end


        # Child elements: pPr, r, fldSimple, hlink, subDoc
        # http://msdn.microsoft.com/en-us/library/office/ee364458(v=office.11).aspx
        def initialize(node, document_properties = {}, doc = nil)
          @node = node
          @properties_tag = 'pPr'
          @document_properties = document_properties
          @font_size = @document_properties[:font_size]
          @document = doc
        end

        # Set text of paragraph
        def text=(content)
          if text_runs.size == 1
            text_runs.first.text = content
          elsif text_runs.size == 0
            new_r = TextRun.create_within(self)
            new_r.text = content
          else
            text_runs.each {|r| r.node.remove }
            new_r = TextRun.create_within(self)
            new_r.text = content
          end
        end

        # Return text of paragraph
        def to_s
          text_runs.map(&:text).join('')
        end

        # Return paragraph as a <p></p> HTML fragment with formatting based on properties.
        def to_html
          html = ''
          text_runs.each do |text_run|
            html << text_run.to_html
          end
          styles = { 'font-size' => "#{font_size}pt" }
          styles['color'] = "##{font_color}" if font_color
          styles['text-align'] = alignment if alignment
          html_tag(:p, content: html, styles: styles)
        end


        # Array of text runs contained within paragraph
        def text_runs
          @node.xpath('w:r|w:hyperlink').map { |r_node| Containers::TextRun.new(r_node, @document_properties) }
        end

        # Iterate over each text run within a paragraph
        def each_text_run
          text_runs.each { |tr| yield(tr) }
        end

        def aligned_left?
          ['left', nil].include?(alignment)
        end

        def aligned_right?
          alignment == 'right'
        end

        def aligned_center?
          alignment == 'center'
        end

        def font_size
          size_attribute = @node.at_xpath('w:pPr//w:sz//@w:val')

          return @font_size unless size_attribute

          size_attribute.value.to_i / 2
        end

        def font_color
          color_tag = @node.xpath('w:r//w:rPr//w:color').first
          color_tag ? color_tag.attributes['val'].value : nil
        end

        def style
          return nil unless @document

          @document.style_name_of(style_id) ||
            @document.default_paragraph_style
        end

        def style_id
          style_property.get_attribute('w:val')
        end

        def style=(identifier)
          id = @document.styles_configuration.style_of(identifier).id

          style_property.set_attribute('w:val', id)
        end

        alias_method :style_id=, :style=
        alias_method :text, :to_s

        # Replace match across all w:t nodes in paragraph order (cross-run).
        # Preserves formatting of the first w:t in the matched span.
        def substitute(match, replacement, multiline: false)
          text_nodes = @node.xpath('.//w:t')
          return if text_nodes.empty?

          merged, node_ranges = build_merged_text(text_nodes)
          matches = collect_substitute_matches(merged, match, replacement)
          return if matches.empty?

          matches.sort_by { |m| -m[:begin] }.each do |m|
            apply_substitute_match(text_nodes, node_ranges, m, multiline: multiline)
          end
        end

        private

        def build_merged_text(text_nodes)
          merged = +''
          ranges = text_nodes.map do |node|
            start = merged.length
            merged << node.content
            { node: node, start: start, end: merged.length }
          end
          [merged, ranges]
        end

        def collect_substitute_matches(merged, match, replacement)
          if match.is_a?(Regexp)
            merged.enum_for(:scan, match).map do
              md = Regexp.last_match
              {
                begin: md.begin(0),
                end: md.end(0),
                replacement: md[0].gsub(match, replacement)
              }
            end
          else
            matches = []
            pos = 0
            while (idx = merged.index(match, pos))
              matches << {
                begin: idx,
                end: idx + match.length,
                replacement: replacement
              }
              pos = idx + match.length
            end
            matches
          end
        end

        def apply_substitute_match(text_nodes, node_ranges, match_info, multiline:)
          match_begin = match_info[:begin]
          match_end = match_info[:end]
          replacement_text = match_info[:replacement]

          affected = node_ranges.select { |nr| nr[:end] > match_begin && nr[:start] < match_end }
          return if affected.empty?

          first_nr = affected.first
          last_nr = affected.last

          prefix_len = match_begin - first_nr[:start]
          prefix = first_nr[:node].content[0, prefix_len]

          suffix_offset = match_end - last_nr[:start]
          suffix = last_nr[:node].content[suffix_offset..-1] || ''

          affected[1..-2].each { |nr| set_wt_content(nr[:node], '') }

          if first_nr.equal?(last_nr)
            write_replacement_to_node(
              first_nr[:node],
              prefix,
              replacement_text,
              suffix,
              multiline: multiline
            )
          else
            write_replacement_to_node(
              first_nr[:node],
              prefix,
              replacement_text,
              nil,
              multiline: multiline
            )
            set_wt_content(last_nr[:node], suffix)
          end
        end

        def write_replacement_to_node(t_node, prefix, replacement_text, suffix, multiline:)
          if multiline && replacement_text.include?("\n")
            write_multiline_replacement(t_node, prefix, replacement_text, suffix)
          else
            set_wt_content(t_node, prefix + replacement_text + (suffix || ''))
          end
        end

        def write_multiline_replacement(t_node, prefix, replacement_text, suffix)
          segments = replacement_text.split("\n", -1)
          set_wt_content(t_node, prefix + segments.first)

          insert_after = t_node
          segments[1..].each do |segment|
            br = Nokogiri::XML::Node.new('w:br', @node.document)
            insert_after.add_next_sibling(br)
            insert_after = br

            new_t = Nokogiri::XML::Node.new('w:t', @node.document)
            set_wt_content(new_t, segment)
            insert_after.add_next_sibling(new_t)
            insert_after = new_t
          end

          if suffix && !suffix.empty?
            set_wt_content(insert_after, insert_after.content + suffix)
          end
        end

        def set_wt_content(t_node, text)
          t_node.content = text
          if text =~ /\A\s|\s\z|  /
            t_node['xml:space'] = 'preserve'
          else
            t_node.remove_attribute('xml:space')
          end
        end

        def style_property
          properties&.at_xpath('w:pStyle') || properties&.add_child('<w:pStyle/>').first
        end

        # Returns the alignment if any, or nil if left
        def alignment
          @node.at_xpath('.//w:jc/@w:val')&.value
        end
      end
    end
  end
end
