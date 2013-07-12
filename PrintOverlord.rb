require 'optparse'
require 'win32/changenotify'
require 'win32ole'
include Win32

@filetypes = ['pdf', 'docx', 'doc', 'html', 'htm', 'xlsx', 'xls']
@home = ENV["HOME"] + "/downloads"
@timeStr = "%m/%d/%Y at %I:%M%p"
@recursion = false

# App title
puts "\n"
puts "-" * 30
puts "Print OverLord"
puts "CTRL + C to stop"
puts "-" * 30
  
# Options
opts = OptionParser.new do |opts|
  opts.banner = "Usage: [flag]"
  
  opts.on( "-f", "--folder STR", String, "Folder to monitor" ) do |v|
    v.gsub!("\\", "/") if v.include? "\\"
    @home = v
    puts "Home is now set to #{v}"
  end
  
  opts.on( "-i", "--inc STR", String, "Include an extension (comma separated, no spaces)" ) do |v|
    @filetypes = @filetypes | v.split(",").map(&:lstrip)
    puts "Extensions are now #{@filetypes.join(",")}"
  end
  
  opts.on( "-e", "--exc STR", String, "Exclude an extension (comma separated, no spaces)" ) do |v|
    @filetypes = @filetypes - v.split(",").map(&:lstrip)
    puts "Extensions are now #{@filetypes.join(",")}"
  end
  
  opts.on( "-o", "--override", "Override extension list with a new list (comma separated, no spaces)" ) do |v|
    @filetypes = v.split(",").map(&:lstrip)
    puts "Extensions are now #{@filetypes.join(",")}"
  end
  
  opts.on( "-r", "--recursion", "Enables folder recursion" ) do |v|
    @recursion = v
    puts "Recursion is now enabled"
  end
  
  opts.on_tail( "-h", "--help", "Show help" ) do
    puts opts
    exit
  end
end.parse!


def main
  drawBanner

  shell = WIN32OLE.new('Shell.Application')
  cn = ChangeNotify.new(@home, @recursion, ChangeNotify::FILE_NAME)
  
  cn.wait do |arr|
    # Loop through events
    arr.each do |info|  
      
      if info.action == "added" || info.action == "renamed new name"
        f = info.file_name.downcase.split(".")

        unless f.last == "tmp" || f.last == "crdownload"  # Skip temporary files
          if @filetypes.include? f.last
            puts "#{Time.now.strftime(@timeStr)} - New file found: #{info.file_name.split("/").last}"
            puts "#{Time.now.strftime(@timeStr)} - Executing Shell PRINT"
            shell.ShellExecute(info.file_name, '', '', 'print', 0)
          else
            puts " Filetype not included: #{f.last}"
          end
          
          drawBanner
        end
        
      end
    end
  end
end

def drawBanner
  puts "\n"
  puts "*" * 30
  puts "#{Time.now.strftime(@timeStr)} - Monitoring for new files in #{@home}"
end

main