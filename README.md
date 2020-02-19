# AudiobookTagger
Powershell script that sorts audiobook files based on folder name and goodreads api. Format &lt;track> - [&lt;series> #&lt;series number>] &lt;title>.  
Scipt takes a folder with multiple folders of audio book tracks and gives them a standard format and the correct info based on GoodReads API.
Point scrpit to author folder (assuming format similar to audiobooks\author\book\)
# Things To Do/Known Issues:
Push info to ID3 file tags (multiple file types)  
Weird Filenames (file names with [] dont get moved)  
Books that use track '0' as intro track  
Sample audio files  
Renames based on original title field. Works the best but not the most effective.(ie. Tiger! Tiger!, The three body problem)  
Good Reads doesn't always have good info. Maybe allow user input for more than just books it can't find.  
Cancel on any out-grid skips folder. (As of now it just messes everything up)  
Same Book multiple audio books  
Same Series but different authors (thinking of leaving this as is but things like the wheel of time, where the last book is written by a different author, might make it harder to find)  
Description cleanup  
Output to csv  
Optimisation (multiple runs out side of the ise cause system slow down. There is a resouce leak or something somewhere)  

# Notes:
Honestly this will problaby the last commit for a powershell version. I'm thinking of switching to Python for many reasons. I started this in POSH because that's what I knew, but there are other directions I want to go with this that I can't do easliy in POSH. GoodReads works well but an overarching issue with audiobooks is that they are not books. There are fields that are different for them (publisher, isbn, ect) that can vary based on which version you have. I'm thinking of using WorldCat. It has an api but it costs. So I'll problaby have to use web scraping. This allong with the actual tag editing would be way easier in python. If anyone actualy reads this, let me know your thoughts.
