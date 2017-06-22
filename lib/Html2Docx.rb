require 'fileutils'
require 'nokogiri'
require 'zip'

require 'Html2Docx/helpers/document_helper'
require 'Html2Docx/helpers/nokogiri_helper'
require 'Html2Docx/helpers/zip_file_generator'

require 'Html2Docx/version'
require 'Html2Docx/initialization'
require 'Html2Docx/content_types'
require 'Html2Docx/relation'
require 'Html2Docx/document'
require 'Html2Docx/document_objects/paragraph'

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

    content_types  = ContentTypes.new(options)

    options[:main_relation] = true
    relation       = Relation.new(options)
    options[:main_relation] = false

    document       = Document.new(options)

    # Render
    document.render
    content_types.render
    relation.render

    # Create Docx File
    self.create_docx(options.fetch(:output), options.fetch(:temp))
  end
end