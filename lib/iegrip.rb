#!ruby -Ks

# IEGrip Ver.0.00    2014/10/06
# Copyright (C) 2014 Yac <iegrip@tech-notes.dyndns.org>
# This software is released under the MIT License, see LICENSE.txt.
require 'win32ole'
require 'win32api'
require 'singleton'
require "iegrip/version"
require 'iegrip/GripWrapper'

module IEgrip
  # ========================
  # IE Application
  # ========================
  class IE < GripWrapper
    def initialize()
      @raw_object = WIN32OLE.new("InternetExplorer.Application")
      @raw_object.visible = true
      fs = FileSystemObject.instance
      ver = fs.GetFileVersion(@raw_object.FullName)
      @majorVersion = ver.split(/\./)[0].to_i
      @urlDownloadToFile = Win32API.new('urlmon', 'URLDownloadToFileA', %w(l p p l l), 'l')
    end
    
    def version
      @majorVersion
    end
    
    def document()
      doc = Document.new(@raw_object.Document, self)
    end
    
    def navigate(url)
      @raw_object.navigate(url)
      sleep 0.5
      wait_stable()
    end
    
    COMPLETE_STATE = 4
    def wait_stable()
      stable_counter = 0
      loop do
        break if stable_counter >= 3
        if (@raw_object.Busy != true) and (@raw_object.ReadyState == COMPLETE_STATE)
          stable_counter += 1
        else
          sleep 0.5
          stable_counter = 0
        end
      end
    end
    
    def export(href, filename)
      @urlDownloadToFile.call(0, href, filename, 0, 0)
    end
  end
  
  module ElementParent
    def parentNode
      raw_element = @raw_object.parentNode()
      raw_element ? HTMLElement.new(raw_element, @ie_obj) : nil
    end
    
    def parentElement
      raw_parent = @raw_object.parentElement
      raw_parent ? HTMLElement.new(raw_parent, @ie_obj) : nil
    end
    
    def getParentForm()
      puts "getParentForm() is called."
      parent_tag = self.parentElement
      loop do
        puts "parent_tag = #{parent_tag.inspect}"
        if parent_tag == nil
          return nil
        elsif parent_tag.tagName == "form"
          return parent_tag
        else
          parent_tag = parent_tag.parentElement
        end
      end
    end
  end
  
  module ElementChild
    def childNodes
      raw_childNodes = @raw_object.childNodes
      raw_childNodes ? TagElementCollection.new(raw_childNodes, @ie_obj) : nil
    end
    
    def previousSibling
      raw_node = @raw_object.previousSibling()
      raw_node ? HTMLElement.new(raw_node, @ie_obj) : nil
    end
    
    def nextSibling
      raw_node = @raw_object.nextSibling()
      raw_node ? HTMLElement.new(raw_node, @ie_obj) : nil
    end
    
    def firstChild
      raw_node = @raw_object.firstChild()
      raw_node ? HTMLElement.new(raw_node, @ie_obj) : nil
    end
    
    def lastChild
      raw_node = @raw_object.lastChild()
      raw_node ? HTMLElement.new(raw_node, @ie_obj) : nil
    end
    
    def hasChildNodes()
      @raw_object.childNodes.each {|subnode|
        return true if (subnode.nodeType != 3) and (subnode.nodeType != 8)
      }
      false
    end
    
    def contains(node)
      @raw_object.contains(toRaw(node))
    end
    
    def isEqualNode(node)
      @raw_object.isEqualNode(toRaw(node))
    end
    
    def Structure(level=0)
      struct = []
      self.childNodes.each {|subnode|
        inner,outer = get_inner(subnode)
        if subnode.hasChildNodes()
          sub_struct = subnode.Structure(level+1)
          if sub_struct.size > 0
            struct.push ("  " * level) + "<#{inner}>"
            struct += sub_struct
            struct.push ("  " * level) + "</#{subnode.tagName}>"
          else
            struct.push ("  " * level) + "<#{inner} />"
          end
        else
          if outer
            struct.push ("  " * level) + "<#{inner}>#{outer}</#{subnode.tagName}>"
          else
            struct.push ("  " * level) + "<#{inner} />"
          end
        end
      }
      return struct
    end
    
    private
    
    def get_inner(tag)
      inner = [tag.tagName]
      outer = nil
      inner.push "id='#{tag.ID}'" if tag.ID != ""
      case tag.tagName
      when "a"
        href = tag.href
        if href.size > 20
          href = href[0,19] + "..."
        end
        inner.push "href='#{href}'"
      when "img"
        inner.push "src='#{tag.src}'"
      when "input"
        inner.push "type='#{tag.Type}'"
      when "form"
        inner.push "action='#{tag.action}' method='#{tag.Method}'"
      when "option"
        inner.push "value='#{tag.value}'"
      when "style"
        inner.push "type='#{tag.Type}'"
      end
      unless tag.hasChildNodes
        innerText = tag.innerText
        if innerText =~ /^<!--(.+)-->$/
          if $1.size > 20
            outer = "<!--#{$1[0,19]}...-->"
          else
            outer = innerText
          end
          innerText = ""
        end
        if innerText.size > 20
          innerText = innerText[0,19] + "..."
        end
        inner.push "text='#{innerText}'" if innerText != ""
      end
      return [inner.join(' '), outer]
    end
  end
  
  module GetElements
    def getElementById(tag_id)
      raw_element = @raw_object.getElementById(tag_id)
      raw_element ? HTMLElement.new(raw_element, @ie_obj) : nil
    end
    
    def getElementsByName(name)
      raw_col = @raw_object.getElementsByName(name)
      raw_col ? TagElementCollection.new(raw_col, @ie_obj) : nil
    end
    
    
    def getElementsByTagName(tag_name)
      raw_col = @raw_object.getElementsByTagName(tag_name)
      raw_col ? TagElementCollection.new(raw_col, @ie_obj) : nil
    end
    alias tags getElementsByTagName
    
    def getTagsByTitle(target_str)
      get_tags_by_key(target_str, "VALUE")
    end
    def getTagsByValue(target_str)
      get_tags_by_key(target_str, "VALUE")
    end
    def getTagsByText(target_str)
      get_tags_by_key(target_str, "INNERTEXT")
    end
    def getTagsByName(target_str)
      get_tags_by_key(target_str, "NAME")
    end
    
    def getTagByTitle(target_str)
      taglist = get_tags_by_key(target_str, "VALUE")
      taglist[0]
    end
    def getTagByValue(target_str)
      taglist = get_tags_by_key(target_str, "VALUE")
      taglist[0]
    end
    def getTagByText(target_str)
      taglist = get_tags_by_key(target_str, "INNERTEXT")
      taglist[0]
    end
    def getTagByName(target_str)
      taglist = get_tags_by_key(target_str, "NAME")
      taglist[0]
    end
    
    private
    
    def get_tags_by_key(target_str, key_type)
      tag_list = []
      @raw_object.all.each {|tag_element|
        case key_type
        when "INNERTEXT"
          key_string = tag_element.innerText
        when "VALUE"
          key_string = tag_element.value
        when "NAME"
          key_string = tag_element.name
        when "ID"
          key_string = tag_element.ID
        else
          return nil
        end
        if key_string == target_str
          tag_list.push HTMLElement.new(tag_element, @ie_obj)
        end
      }
      
      return tag_list
    end
  end
  
  # ========================
  # Node
  # ========================
  class Node < GripWrapper
    NODETYPE_DIC = {
      1 => :ELEMENT_NODE,
      2 => :ATTRIBUTE_NODE,
      3 => :TEXT_NODE,
      4 => :CDATA_SECTION_NODE,
      5 => :ENTITY_REFERENCE_NODE,
      6 => :ENTITY_NODE,
      7 => :PROCESSING_INSTRUCTION_NODE,
      8 => :COMMENT_NODE,
      9 => :DOCUMENT_NODE,
      10 => :DOCUMENT_TYPE_NODE,
      11 => :DOCUMENT_FRAGMENT_NODE,
      12 => :NOTATION_NODE,
    }
    
    def nodeName
      @raw_object.nodeName
    end
    
    def nodeType
      @raw_object.nodetype
    end
    
    def nodeTypeName
      nodetype = @raw_object.nodetype
      NODETYPE_DIC[nodetype] || :UNKNOWN
    end
    
    def inspect
      "<#{self.class}, Name:#{self.nodeName}>"
    end
    
  end
  
  # ========================
  # IE.Document
  # ========================
  class Document  < Node
    include ElementChild
    include GetElements
    
    def head()
      raw_head = @raw_object.head
      raw_head ? HTMLElement.new(raw_head, @ie_obj) : nil
    end
    
    def body()
      HTMLElement.new(@raw_object.body, @ie_obj)
    end
    
    def all
      raw_all = @raw_object.all
      raw_all ? TagElementCollection.new(raw_all, @ie_obj) : nil
    end
    
    def frames(index=nil)
      if index
        return(nil) if index >= @raw_object.Frames.length
        Frames.new(@raw_object.frames, @ie_obj)[index]
      else
        Frames.new(@raw_object.frames, @ie_obj)
      end
    end
    
    def documentElement
      raw_element = @raw_object.documentElement()
      raw_element ? HTMLElement.new(raw_element, @ie_obj) : nil
    end
    
    def createElement(tag_name)
      raw_element = @raw_object.createElement(tag_name)
      HTMLElement.new(raw_element, @ie_obj)
    end
    
    def createAttribute(attr_name)
      raw_attr = @raw_object.createAttribute(attr_name);
      Attr.new(raw_attr, @ie_obj)
    end
    
  end
  
  # ========================
  # TAG Element
  # ========================
  class HTMLElement  < Node
    include ElementParent
    include ElementChild
    include GetElements
    def tagname
      if self.nodeType == 8
        "comment"
      else
        @raw_object.tagName.downcase
      end
    end
    
    def text=(set_text)
      case self.tagname
      when "select"
        option_list = tags("OPTION")
        option_list.each {|option_element|
          if option_element.innerText == set_text
            option_element.selected = true
            break
          end
        }
      else
        @raw_object.value = set_text
      end
    end
    
    def inspect()
      case tagName
      when "SELECT"
        innerHTML = replace_cr_code(self.innerHTML)
        "<#{self.class}, TAG:#{tagName}, [#{self.innerHTML}]"
      when "INPUT", "IMG", "A"
        outerHTML = replace_cr_code(self.outerHTML)
        "<#{self.class}, TAG:#{tagName}, [#{self.outerHTML}]"
      when "TR", "TD"
        innerText = replace_cr_code(self.innerText)
        "<#{self.class}, TAG:#{tagName}, [#{innerText}]"
      else
        "<#{self.class}, TAG:#{tagName}>"
      end
    end
    
    def to_s
      @raw_object.value
    end
    def value
      @raw_object.value
    end
    alias text value
    
    def click
      if @ie_obj.version >= 10
        case self.tagname.downcase
        when "a"
          href = self.href
          @ie_obj.navigate(href)
        when "input"
          if self.Type.downcase == "submit"
            puts "**** Submit Type is detected."
            parent_form = self.getParentForm()
            if parent_form
              puts "parent_form = #{parent_form.outerHTML}"
              ret_val = parent_form.submit()
              puts "parent_form.submit() submit is called. ret_val = #{ret_val.inspect}"
              sleep 1
              parent_form.raw.submit()
              parent_form.fireEvent("onSubmit")
              parent_form.fireEvent("onClick")
            else
              puts "parent_form not detected."
            end
          end
        when "button"
          @raw_object.fireEvent("onClick")
        else
          @raw_object.click
        end
      else
        @raw_object.click
      end
      @ie_obj.wait_stable()
    end
    
    def all
      TagElementCollection.new(@raw_object.all, @ie_obj)
    end
    
    
    def export(filename)
      case self.tagName.downcase
      when "img"
        @ie_obj.export(self.src, filename)
      when "a"
        @ie_obj.export(self.href, filename)
      else
        raise "export() is not support."
      end
    end
    
    def setAttributeNode(attribute)
      @raw_object.setAttributeNode(toRaw(attribute));
    end
    
    def getAttribute(attr_name)
      @raw_object.getAttribute(attr_name)
    end
    
    def getAttributeNode(attr_name)
      raw_attr = @raw_object.getAttributeNode(attr_name)
      raw_attr ? Attr.new(raw_attr, @ie_obj) : nil
    end
    
    def removeAttribute(attr_name)
      @raw_object.removeAttribute(attr_name)
    end
    
    def removeAttributeNode( attr )
      raw_attr = @raw_object.removeAttributeNode( toRaw(attr) )
      raw_attr ? Attr.new(raw_attr, @ie_obj) : nil
    end
    
    private
    
    def replace_cr_code(text)
      replcae_text = text.gsub(/\r/, '\r')
      replcae_text.gsub!(/\n/, '\n')
      return replcae_text
    end
  end
  
  # ========================
  # TAG Element Collection
  # ========================
  class TagElementCollection  < GripWrapper
    def [](index)
      return(nil) if index >= @raw_object.length
      HTMLElement.new(@raw_object.item(index), @ie_obj)
    end
    
    def size
      @raw_object.length
    end
    
    def each
      @raw_object.each {|tag_element|
        next if (tag_element.nodeType == 3) or (tag_element.nodeType == 8)
        yield HTMLElement.new(tag_element, @ie_obj)
      }
    end
    
    def getTagsByTitle(target_str)
      get_tags_by_key(target_str, "VALUE")
    end
    def getTagsByValue(target_str)
      get_tags_by_key(target_str, "VALUE")
    end
    def getTagsByText(target_str)
      get_tags_by_key(target_str, "INNERTEXT")
    end
    def getTagsByName(target_str)
      get_tags_by_key(target_str, "NAME")
    end
    
    def getTagByTitle(target_str)
      taglist = get_tags_by_key(target_str, "VALUE")
      taglist ? taglist[0]: nil
    end
    def getTagByValue(target_str)
      taglist = get_tags_by_key(target_str, "VALUE")
      taglist ? taglist[0]: nil
    end
    def getTagByText(target_str)
      taglist = get_tags_by_key(target_str, "INNERTEXT")
      taglist ? taglist[0]: nil
    end
    def getTagByName(target_str)
      taglist = get_tags_by_key(target_str, "NAME")
      taglist ? taglist[0]: nil
    end
    
    def inspect()
      tagname_list = []
      self.each {|tag_element|
        tagname_list.push "<#{tag_element.tagName}>"
      }
      if tagname_list.size > 3
        "<#{self.class}: [#{tagname_list[0,3].join(', ')},...]"
      else
        "<#{self.class}: [#{tagname_list.join(', ')}]>"
      end
    end
    
    private
    
    def get_tags_by_key(target_str, key_type)
      tag_list = []
      @raw_object.each {|tag_element|
        case key_type
        when "INNERTEXT"
          key_string = tag_element.innerText
        when "VALUE"
          key_string = tag_element.value
        when "NAME"
          key_string = tag_element.name
        when "ID"
          key_string = tag_element.ID
        else
          return nil
        end
        if key_string == target_str
          tag_list.push HTMLElement.new(tag_element, @ie_obj)
        end
      }
      case tag_list.size
      when 0
        return nil
      else
        return tag_list
      end
    end
    
  end


  
  # ========================
  # IE.Document.Frames
  # ========================
  class Frames  < GripWrapper
    def [](index)
      return(nil) if index >= @raw_object.length
      Frame.new(@raw_object.item(index), @ie_obj)
    end
    
    def size
      @raw_object.length
    end
    
    def each
      index = 0
      while index < @raw_object.length
        raw_frame = @raw_object.item(index)
        yield Frame.new(raw_frame, @ie_obj)
        index += 1
      end
    end
  end
  
  # ========================
  # IE.Document.Frames.item(n)
  # ========================
  class Frame  < GripWrapper
    def document
      Document.new(@raw_object.document, @ie_obj)
    end
  end
  
  class Attr < GripWrapper
    def value=(value_str)
      @raw_object.value = value_str
    end
    def value
      @raw_object.value
    end
    
    def ownerElement()
      HTMLElement.new(@raw_object.ownerElement, @ie_obj)
    end
  end
end



