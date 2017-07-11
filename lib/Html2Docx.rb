require 'fileutils'
require 'nokogiri'
require 'zip'
require 'typhoeus'

require 'Html2Docx/helpers/document_helper'
require 'Html2Docx/helpers/nokogiri_helper'
require 'Html2Docx/helpers/zip_file_generator'

require 'Html2Docx/version'
require 'Html2Docx/initialization'
require 'Html2Docx/content_types'
require 'Html2Docx/relation'
require 'Html2Docx/document'

require 'Html2Docx/document_objects/paragraph'
require 'Html2Docx/document_objects/heading'
require 'Html2Docx/document_objects/image'

module Html2Docx
  ROOT_PATH = File.expand_path(File.join(File.dirname(__FILE__), '../'))

  def self.clear_temp(tmp)
    FileUtils.rm_r tmp
  end

  def self.create_docx(output, input)
    zf = ZipFileGenerator.new(input, output)
    zf.write

    self.clear_temp(input)
  end

  def self.render(options = {})
    initialization = Initialization.new(options)
    options[:temp] = initialization.get_temp_directory
    options[:initialization] = initialization

    content_types  = ContentTypes.new(options)
    options[:content_types] = content_types

    options[:main_relation] = true
    main_relation  = Relation.new(options)
    options[:main_relation] = main_relation

    document       = Document.new(options)
    options[:document] = document

    # Render
    document.render
    content_types.render
    main_relation.render

    # Create Docx File
    self.create_docx(options.fetch(:output), options.fetch(:temp))
  end
end