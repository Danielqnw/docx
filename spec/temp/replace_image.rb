#!/usr/bin/env ruby
# frozen_string_literal: true

require 'docx'
require 'fileutils'

docx_path = ARGV[0]
image_selector = ARGV[1]
replacement_path = ARGV[2]
output_path = ARGV[3] || 'tmp/replaced.docx'

if [docx_path, image_selector, replacement_path].any? { |arg| arg.nil? || arg.strip.empty? }
  warn "Usage: ruby spec/temp/replace_image.rb <docx_path> <rid|entry_path|index> <replacement_image_path> [output_docx_path]"
  exit 1
end

unless File.exist?(docx_path)
  warn "DOCX not found: #{docx_path}"
  exit 1
end

unless File.exist?(replacement_path)
  warn "Replacement image not found: #{replacement_path}"
  exit 1
end

doc = Docx::Document.open(docx_path)
images = doc.images

if images.empty?
  warn "No images found in: #{docx_path}"
  exit 1
end

reference =
  if image_selector.match?(/^\d+$/)
    index = image_selector.to_i
    if index <= 0 || index > images.size
      warn "Index out of range: #{index}. Valid range is 1..#{images.size}"
      exit 1
    end
    images.to_a[index - 1][1]
  else
    image_selector
  end

doc.replace_image(reference, replacement_path)

output_dir = File.dirname(output_path)
FileUtils.mkdir_p(output_dir) unless output_dir == '.'

doc.save(output_path)
puts "Done: replaced #{reference} -> #{replacement_path}"
puts "Saved: #{output_path}"
