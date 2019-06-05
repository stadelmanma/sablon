module Sablon
  # Combines the user supplied context and template into a single object
  # to manage data during template processing.
  class Environment
    attr_reader :document
    attr_reader :context
    attr_reader :section_properties

    # returns a new environment with merged contexts
    def alter_context(context = {})
      new_context = @context.merge(context)
      Environment.new(document, new_context)
    end

    def section_properties=(properties)
      @section_properties = Context.transform_hash(properties)
    end

    private

    def initialize(document, context = {})
      @document = document
      @context = Context.transform_hash(context)
      @section_properties = {}
    end
  end
end
