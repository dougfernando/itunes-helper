require 'win32ole'
require 'soap/wsdlDriver' 
require 'iconv'

class ITunesHelper
  def initialize(lyrics_service)
    @itunes = WIN32OLE.new('iTunes.Application')
    @tracks = @itunes.LibraryPlaylist.Tracks
    @lyrics_service = lyrics_service
  end

  def show_all_tracks
    @tracks.each do |track|
      print track.Name() + "\n"
    end
  end

  def set_all_lyrics
    puts "Number of tracks to be analyzed: #{@tracks.count}"
    index = 1
    @tracks.each do |track|
      if self.is_lyrics_empty(track.lyrics)
        lyrics = @lyrics_service.get_lyrics(track.Artist(), track.Name())
        
        if self.is_lyrics_empty(lyrics)
          self.print_track(index, @tracks.count, track, "lyrics WAS NOT found")
        else
          track.lyrics = @lyrics_service.get_lyrics(track.Artist(), track.Name())
          self.print_track(index, @tracks.count, track, "lyrics WAS found and set")
        end
      else
        self.print_track(index, @tracks.count, track, "already had lyrics")
      end
      index = index + 1
    end
  end

  def is_lyrics_empty(lyrics)
    lyrics.empty? or lyrics == "Not found"
  end

  def print_track(index, total, track, result) 
    puts "Track #{index}/#{total}: #{track.Artist()} - #{track.Name()}: #{result}"
  end
end

class LyricsServiceProxy
  def initialize
    @driver = SOAP::WSDLDriverFactory.new("http://lyricwiki.org/server.php?wsdl").create_rpc_driver
    puts "Lyrics web service initialized"
  end

  def get_lyrics(artist, track_name)
    Iconv.iconv("LATIN1", "UTF-8", @driver.getSong(artist, track_name).lyrics).to_s
  end
end

itunesH = ITunesHelper.new(LyricsServiceProxy.new)
itunesH.set_all_lyrics
