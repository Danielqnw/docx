require 'docx/version'

module Docx #:nodoc:
  autoload :Document, 'docx/document'

  module Merge
    autoload :StylesImporter, 'docx/merge/styles_importer'
  end
end

