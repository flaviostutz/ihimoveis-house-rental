# ihimoveis-house-rental

This was the first online house rental website in Brazil, launched on 2001!

I developed it along with Hugo Gonzalez (he was the designer) for the rental agency of his father, who at the time bought a digital camera (the ones that uses spinning disks) and had the vision to use these photos for creating a digital portfolio of the houses at rental, so we, as teenagers, engaged in this project!

This website was run with this system for various years.

If you are curious, check the website at [Wayback Machine in 2002!](https://web.archive.org/web/20020515000000*/ihimoveis.com.br)

<img src="screenshot.png" with="600"></img>

## Constraints at the time

* Internet connection was very expensive and we had to upload photos during the night, because the rates were much lower

* The advertisements should be prepared during the day, when employees were at the office


## Features of the system

* During the day, the employees used a Windows machine with a local running IIS server that they used to access the web site and prepare the advertisements with description and photos

* During the night there was a cron script in the Windows machine that optimized the image sizes (jpg compression and thumbnails creation), then uploaded the photos and the database to the server via a ftp service

* We could see which houses had the most visits also 💪

## Remarks

At the time the agency won some prices and appeared in newspapers because of this solution. It was very nice to see that their business changed a lot after employing the website, for example, because they started reaching people from other cities, which was very hard before as they told us!

It was a lot of fun doing this!!!
