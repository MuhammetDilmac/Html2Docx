module Html2Docx
  class Relation
    def initialize(options = {})
      @relation_file = nil
      @relation = nil
      @relations = []
      @last_relation_id = 1

      if options[:main_relation]
        @relation_file = File.join(TEMP_PATH, 'word', '_rels', 'document2.xml.rels')
        @relation = File.open(@relation_file) { |f| Nokogiri::XML(f) }
        @last_relation_id = @relation.css('Relationship').last.attr('Id').to_i
      else
        @relation_file = File.join(TEMP_PATH, 'word', '_rels', options.fetch(:file_name))
        @relation = create_relation_file
      end

      @relations = @relation.css('Relationship').first
    end

    def create_relation_file
      document = Nokogiri::XML::Document.new
      document.encoding = 'UTF-8'
      relations_tag = Nokogiri::XML::Node.new('Relationships', document)
      relations_tag['xmlns'] = 'http://schemas.openxmlformats.org/package/2006/relationships'
      document.add_child relations_tag
      document
    end

    def add_relation(type, target)
      relation_tag = Nokogiri::XML::Node.new('Relationship', @relation)
      relation_tag['Id'] = "rId#{@last_relation_id + 1}"
      relation_tag['Type'] = get_type(type)
      relation_tag['Target'] = get_target(target)
      @relations.add_child(relation_tag)
      @last_relation_id = @last_relation_id + 1
    end

    def render
       File.open(@relation_file, 'w') { |f| f.write(Helpers::NokogiriHelper.to_xml(@relation)) }
    end

    def get_type(type)
    end

    def get_target(target)
    end
  end
end