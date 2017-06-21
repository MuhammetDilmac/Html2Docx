require 'fileutils'
require 'nokogiri'

require 'Html2Docx/helpers/document_helper'
require 'Html2Docx/helpers/nokogiri_helper'

require 'Html2Docx/version'
require 'Html2Docx/initialization'
require 'Html2Docx/content_types'
require 'Html2Docx/relation'
require 'Html2Docx/document'
require 'Html2Docx/document_objects/paragraph'

module Html2Docx
  ROOT_PATH = File.expand_path(File.join(File.dirname(__FILE__), '../'))
  TEMP_PATH = Dir.mktmpdir

  def self.start(options = {})
    initialization = Initialization.new(options)

    content_types  = ContentTypes.new(options)

    options[:main_relation] = true
    relation       = Relation.new(options)

    options[:html] = <<-HTML
    <!DOCTYPE html>
<html>
<body>
  <p style="text-align: right;"><strong>Lorem</strong> <i>Ipsum</i>, <u>dizgi</u> <s>ve</s> <font color="#ff0000">baskı</font> endüstrisinde kullanılan mıgır metinlerdir.</p>
  <p style="text-align: justify;">Lorem Ipsum, adı bilinmeyen bir matbaacının bir hurufat numune kitabı oluşturmak üzere bir yazı galerisini alarak karıştırdığı 1500'lerden beri endüstri standardı sahte metinler olarak kullanılmıştır.</p>
  <p style="background-color: #00ff00;">Beşyüz yıl boyunca varlığını sürdürmekle kalmamış, aynı zamanda pek değişmeden elektronik dizgiye de sıçramıştır.</p>
  <p style="text-indent: 5px;">1960'larda Lorem Ipsum pasajları da içeren Letraset yapraklarının yayınlanması ile ve yakın zamanda Aldus PageMaker gibi Lorem Ipsum sürümleri içeren masaüstü yayıncılık yazılımları ile popüler olmuştur.</p>
  <p style="text-indent: 30px;">Yaygın inancın tersine, Lorem Ipsum rastgele sözcüklerden oluşmaz.</p>
  <p style="text-align: center; line-height: 2; background-color: #0000ff">Kökleri M.Ö. 45 tarihinden bu yana klasik Latin edebiyatına kadar uzanan 2000 yıllık bir geçmişi vardır. </p>
</body>
</html>
    HTML
    document       = Document.new(options)

    # Render
    document.render
    content_types.render
    relation.render
  end

  def self.temp_path
    TEMP_PATH
  end

  def self.root_path
    ROOT_PATH
  end

  def self.clear_temp
    FileUtils.rm_r TEMP_PATH
  end

  def self.get_file(value)
    file_path = File.join(TEMP_PATH, value)
    file_content = File.read(file_path)

    puts file_content
  end
end

Html2Docx.start