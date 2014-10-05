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
    
    def navigate(url, flags=nil, targetFrameName=nil, postData=nil, headers=nil)
      @raw_object.navigate(url, flags, targetFrameName, postData, headers)
      wait_stable()
    end
    
    COMPLETE_STATE = 4
    def wait_stable()
      while @raw_object.Busy == true
        sleep 0.5
      end
      loop do
        break if @raw_object.ReadyState == COMPLETE_STATE
        sleep 0.5
      end
    end
    
    def export(href, filename)
      @urlDownloadToFile.call(0, href, filename, 0, 0)
    end
  end
  
  # ========================
  # IE.Document
  # ========================
  class Document  < GripWrapper
    def frames(index=nil)
      if index
        return(nil) if index >= @raw_object.Frames.length
        Frames.new(@raw_object.frames, @ie_obj)[index]
      else
        Frames.new(@raw_object.frames, @ie_obj)
      end
    end
    
    def all()
      TagElementCollection.new(@raw_object.all, @ie_obj)
    end
    
    def tags(tag_name)
      raw_col = @raw_object.all.tags(tag_name)
      raw_col ? TagElementCollection.new(raw_col, @ie_obj) : nil
    end
    
    def head()
      raw_head = @raw_object.head
      raw_head ? TagElement.new(raw_head, @ie_obj) : nil
    end
    
    def body()
      raw_body = @raw_object.body
      raw_body ? TagElement.new(raw_body, @ie_obj) : nil
    end
    
    def getElementById(tag_id)
      raw_element = @raw_object.getElementById(tag_id)
      raw_element ? TagElement.new(raw_element, @ie_obj) : nil
    end
    
    def getElementsByName(tag_name)
      raw_col = @raw_object.getElementsByName(tag_name)
      raw_col ? TagElementCollection.new(raw_col, @ie_obj) : nil
    end
    
    def getElementsByTagNameNS(pvarNS, bstrLocalName)
      raw_col = @raw_object.getElementsByTagNameNS(pvarNS, bstrLocalName)
      raw_col ? TagElementCollection.new(raw_col, @ie_obj) : nil
    end
  end
  
  
  # ========================
  # TAG Element Collection
  # ========================
  class TagElementCollection  < GripWrapper
    def [](index)
      return(nil) if index >= @raw_object.length
      TagElement.new(@raw_object.item(index), @ie_obj)
    end
    
    def size
      @raw_object.length
    end
    
    def each
      @raw_object.each {|tag_element|
        yield TagElement.new(tag_element, @ie_obj)
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
    def getTagByID(target_id)
      taglist = get_tags_by_key(target_id, "ID")
      taglist ? taglist[0]: nil
    end
    
    def inspect()
      element_inspect_list = []
      self.each {|tag_element|
        element_inspect_list.push tag_element.inspect
      }
      "#{self.class},#{element_inspect_list.join('/')}"
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
          tag_list.push TagElement.new(tag_element, @ie_obj)
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
  # TAG Element
  # ========================
  class TagElement  < GripWrapper
    def tagname
      @raw_object.tagName.downcase
    end
    
    def text=(set_text)
      case tagName
      when "SELECT"
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
        "#{self.class}:<#{tagName}>:[#{self.innerHTML}]"
      when "INPUT", "IMG", "A"
        outerHTML = replace_cr_code(self.outerHTML)
        "#{self.class}:<#{tagName}>:[#{self.outerHTML}]"
      when "TR", "TD"
        innerText = replace_cr_code(self.innerText)
        "#{self.class}:<#{tagName}>:[#{innerText}]"
      else
        "#{self.class}:<#{tagName}>"
      end
    end
    
    def to_s
      @raw_object.value
    end
    
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
    
    def tags(tag_name)
      #raw_col = @raw_object.getElementsByName(tag_name)
      raw_col = @raw_object.all.tags(tag_name)
      TagElementCollection.new(raw_col, @ie_obj)
    end
    
    def all
      TagElementCollection.new(@raw_object.all, @ie_obj)
    end
    
    def parentElement
      raw_parent = @raw_object.parentElement
      raw_parent ? TagElement.new(raw_parent, @ie_object) : nil
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
    
    private
    
    def replace_cr_code(text)
      replcae_text = text.gsub(/\r/, '\r')
      replcae_text.gsub!(/\n/, '\n')
      return replcae_text
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
        ie_frame = Frame.new(raw_frame, @ie_obj)
        yield(ie_frame)
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
end



