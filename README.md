# IMDbScraper

IMDbScraper, .NET Dynamic-link library for scraping movie information from IMDb.

The library written under concern of Movie, might break if used against TV series.  
The is a PROOF OF CONCEPT on automate data retrieval from IMDb.com  
It's fairly easy to integrated into other VB project since it's a managed DLL but you shouldn't use it, especially as data grabbing tools.

If interested on movies data in Chinese, kindly check out it's sister project [MtimeScraper](https://github.com/hebeguess/MtimeScraper) (Chinese variant of IMDb).

**NOTICE** : It's been a while since my last contribution on the project, some parts of it may be broken.


### Usage
-----
* Load and compiled the project in Visual Studio.
* Import the generated DLL into your desire projects.  

Declare new IMDb Class, pass either html source OR movie title to it's constructor.

    // While you already have html content
    IMDb i = new IMDb(string pageSource);

    // Library will handle content loading for you; movie numbers from URL, OPTIONAL timeout value for WebRequest
    IMDb i = new IMDb(Integer movNo, Integer timeOut);

    // Library usage is simple, IDE will assists you along the way. Just try out yourself.
    string movieTitle = i.getTitle();

Remember to apply threading if you're using the latter to avoid deadlock while waiting web response complete.


### Terms of Use
------------

IMDb or its content providers does not permit use of its data by third parties without their consent.

The project was motivated solely under personal purposed.

Please note you SHOULD NOT using this library from anything other than limited, personal and non-commercial use.

Neither I, nor any subsequent contributors of the project hold any kind of liability caused by your usage.
