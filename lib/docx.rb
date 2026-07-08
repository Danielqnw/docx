require 'docx/version'

module Docx #:nodoc:
  autoload :Document, 'docx/document'

  module Merge
    autoload :StylesImporter, 'docx/merge/styles_importer'
    autoload :NumberingImporter, 'docx/merge/numbering_importer'
    autoload :NodeRewriter, 'docx/merge/node_rewriter'
    autoload :Importer, 'docx/merge/importer'
  end
end

