<#
.SYNOPSIS
  Name: query-movie.ps1
  The purpose of this script is to scrape to web.
  
.DESCRIPTION
  I was looking for an automated way to pull movie info based on a query.  Using PowerShell's web scraping cmdlets I created a prototype to scrap from movie databases.
  
.NOTES
    Initial Release: 2018-18-01
   
  Author: Ben Beard (@capt_beard)

.EXAMPLE
  Run the query-movie script and enter a movie title when prompted.
  If desired, results will be exported to CSV file in your temp directory.

#>

do { #Loop through entire script

#Get info from user
$title = read-host "Enter the name of the movie you wish to search for"

#Convert title to searcy query string
$query = $title -replace '\s',"+"

#Set variable for search link
$url = "https://www.imdb.com/find?s=all&q=$query"

#Pull htm from url variable
$search = invoke-webrequest -uri $url

#Serach web page for specific class, select first 10 results, pop-up selection box, select property based on user selection
$imdb = $search.AllElements | where Class -eq "result_text" | select -first 10 | out-gridview -title "Pick the appropriate title" -passthru | select -expandproperty innerhtml 

#Get IMDB number from link
$imdbnum = $imdb.Split("/title/", 9) | select -last 1
$imdbnum = $imdbnum.Split("/ref") | select -first 1

#Set IMDB link variable
$imdblink = "https://www.imdb.com/title/$imdbnum"

#Scrape web based on IMDB link
$newsearch = invoke-webrequest -uri $imdblink

#Search page for specific class, select first result
$movie = $newsearch.AllElements | where Class -eq "subtext" | select -first 1 -expandproperty innertext

#Extract specific information from selected element
$rating = $movie.split("|") | select -first 1
$genre = $movie.split("|") | select -skip 2 -first 1
$dir = $newsearch.AllElements | where Class -eq "itemprop" | where itemprop -eq "name" | select -first 1 -expandproperty innertext
$title = $newsearch.ParsedHtml.GetElementsByTagName('title') | select -first 1 -expandproperty innertext
$title = $title.Split("-") | select -first 1

#Print obtained information to screen
write-host
write-host "*-------------------------------------------------------*"
write-host "Title: $title"
write-host "Rating: $rating"
write-host "Genre: $genre"
write-host "Director: $dir"
write-host "Results from IMDB"
write-host "*-------------------------------------------------------*"
write-host 

#Give user option to export to CSV
$export = read-host "Would you like this exported to Excel? (yes/no)"
	if($export -eq "yes")
		{
		#Add elements to array
		new-object -typename PSObject -Property @{
			Film = $title
			Rating = $rating
			Genre = $genre
			Director = $dir
		} | select-object Film,Rating,Genre,Director | export-csv C:\temp\imdb.csv -notypeinformation -append #Export array to CSV
		write-host "Your information has been exported to C:\temp\imdb.csv"
		write-host
		}

#Option to loop through entire script		
$ans = read-host "Would you like to run again? (yes/no)"
write-host

}
#Logic to end loop based on user answer
While($ans -eq "yes")


