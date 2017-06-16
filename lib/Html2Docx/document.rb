module Html2Docx
  class Document
    def initialize(options = {})
      @document_file = File.join(TEMP_PATH, 'word', 'document2.xml')
      @document = File.open(@document_file) { |f| Nokogiri::XML(f) }
      @body = @document.css('w:body')
      @contents = []

      initial_body
    end

    def initial_body
      @body.remove
    end

    def add_html(html)
      html = Nokogiri::HTML(html.gsub!(/\sl\s+|\n/, ' '))

      elements = html.css('body')

      elements.children.each do |element|
        case element.name
          when 'p'
            # Add paragraph
            @contents.push ''
          when 'table'
            # Add table
            @contents.push ''
        end
      end
    end

    def render
      @contents.each do |content|
        @body << content
      end

      @body << sectPr

      File.open(@document_file, 'w') { |f| f.write(@document.to_xml) }
    end

    def sectPr
      Nokogiri::XML::Builder.new do |xml|
        xml.send('w:sectPr') do
          xml.send('w:pgSz', { 'w:w' => '12240', 'w:h' => '15840' })
          xml.send('w:pgMar', {
                     'w:top' => '1440', 'w:right' => '1440', 'w:bottom' => '1440', 'w:left' => '1440',
                     'w:header' => '720', 'w:footer' => '720', 'w:gutter' => '0'
                   })
          xml.send('w:cols', { 'w:space' => '720' })
          xml.send('w:docGrid', { 'w:linePitch' => 360 })
        end
      end.root
    end
  end
end