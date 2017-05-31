module Sablon
  # tracks bookmarks for use in creating cross-references
  class Bookmarks
    def initialize
      @counter = 0
      @names = []
    end

    def initialize_bookmark_ids(xml_node)
      @counter = xml_node.xpath('.//w:bookmarkStart').inject(0) do |max, node|
        @names << node['w:name']
        [node['w:id'].to_i, max].max
      end
    end

    # adds a new bookmark name to prevent shadowing later in the doc
    def <<(bookmark)
      if @names.include? bookmark.name
        raise ContextError, "Bookmark name already in use: #{bookmark.name}"
      end
      #
      @names << bookmark.name
      @counter += 1
      bookmark.id = @counter.to_s
    end
  end
end
