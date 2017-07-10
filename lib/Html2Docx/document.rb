module Html2Docx
  class Document
    def initialize(options = {})
      @tmp_path = options[:temp]
      @document_file = File.join(@tmp_path, 'word', 'document2.xml')
      @document = File.open(@document_file) { |f| Nokogiri::XML(f) }
      @body = @document.at_xpath('//w:body')
      @contents = []
      @relation = options[:main_relation]

      initial_body
      add_html(options[:html])
    end

    def initial_body
      @body.children.each do |child|
        child.remove
      end
    end

    def add_html(html)
      html = Nokogiri::HTML(html.gsub!(/\sl\s+|\n/, ' '))

      elements = html.css('body')

      elements.children.each do |element|
        case element.name
          when 'p'
            # Add paragraph
            paragraph = DocumentObjects::Paragraph.new(@document, @relation, @tmp_path)
            paragraph.add_paragraph(element)
            @contents.push paragraph.render
          when /h[1-9]/
            heading = DocumentObjects::Heading.new(@document, @relation, @tmp_path)
            heading.add_heading(element)
            @contents.push heading.render
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

      @document.root.add_child(@body)

      File.open(@document_file, 'w') { |f| f.write(Helpers::NokogiriHelper.to_xml(@document)) }
    end

    def sectPr
      root = Nokogiri::XML::Node.new('w:sectPr', @document)
      root.add_child(pgSz)
      root.add_child(pgMar)
      root.add_child(cols)
      root.add_child(docGrid)

      root
    end

    def pgSz
      pgSz = Nokogiri::XML::Node.new('w:pgSz', @document)
      pgSz['w:w'] = '12240'
      pgSz['w:h'] = '15840'

      pgSz
    end

    def pgMar
      pgMar = Nokogiri::XML::Node.new('w:pgMar', @document)
      pgMar['w:top'] = '1440'
      pgMar['w:right'] = '1440'
      pgMar['w:bottom'] = '1440'
      pgMar['w:left'] = '1440'
      pgMar['w:header'] = '720'
      pgMar['w:footer'] = '720'
      pgMar['w:gutter'] = '0'

      pgMar
    end

    def cols
      cols = Nokogiri::XML::Node.new('w:cols', @document)
      cols['w:space'] = '720'

      cols
    end

    def docGrid
      docGrid = Nokogiri::XML::Node.new('w:docGrid', @document)
      docGrid['w:linePitch'] = '360'

      docGrid
    end
  end
end