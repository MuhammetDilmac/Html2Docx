module Html2Docx
  class Relation
    def initialize(options = {})
      @relation_file = nil
      @relation = nil
      @relations = []
      @last_relation_id = 1
      @internal_links = {}

      if options[:main_relation]
        @relation_file = File.join(options.fetch(:temp), 'word', '_rels', 'document2.xml.rels')
        @relation = File.open(@relation_file) { |f| Nokogiri::XML(f) }
        @last_relation_id = @relation.css('Relationship').last.attr('Id').to_i
      else
        @relation_file = File.join(options.fetch(:temp), 'word', '_rels', options.fetch(:file_name))
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

    def get_type(type)
    end

    def get_target(target)
    end

    def create_internal_link_start_tag(name, document)
      bookmark_start_tag = Nokogiri::XML::Node.new('w:bookmarkStart', document)
      bookmark_start_tag['w:id'] = create_internal_link_id(name)
      bookmark_start_tag['w:name'] = name

      bookmark_start_tag
    end

    def create_internal_link_end_tag(name, document)
      bookmark_end_tag = Nokogiri::XML::Node.new('w:bookmarkEnd', document)
      bookmark_end_tag['w:id'] = find_internal_link_id(name)

      bookmark_end_tag
    end

    def create_internal_link_id(name)
      id = find_internal_link_id(name)
      if id
        id = get_latest_internal_link_id + 1
        @internal_links[id] = name
      else
        id
      end
    end

    def get_latest_internal_link_id
      @internal_links.keys.max || 0
    end

    def find_internal_link_id(name)
      @internal_links.find{ |key, value| value == name }
    end

    def render
       File.open(@relation_file, 'w') { |f| f.write(Helpers::NokogiriHelper.to_xml(@relation)) }
    end
  end
end