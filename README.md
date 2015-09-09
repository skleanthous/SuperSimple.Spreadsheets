SuperSimple.Spreadsheet
=======================

1. Description
--------------

This is a library to make writing spreadsheets using the OpenXML SDK 2.5 as easy as possible. I did not set to create this project (and nuget package, see installation) due to vanity or to learn OpenXML. Rather, two years ago when I first started to had to read and write spreadsheet files, there was no free and open source project to do this in a simple way. 

The goal is to allow people to shift focus away from working with spreadsheets, lift the impediment to learn libraries to read and save spreadsheets, and cut down the time needed to start working with spreadsheets to a minimum.

2. Things to be aware
---------------------

This project has been developed in two working days some time ago. I found it useful, but it direly needs some work. I will put the effort required to publish, create a help page and add some releasable testing in the project. Please use it with the fact that this will possibly break if you try to do anything complex. Stick to simple usage and things will work.

Do not expect features that will make spreadsheets fancier to the eye that will have an impact to usability: this is aimed to provide a spreadsheet with data to the ones in the business to use with minimal effort by the devs. If you need something much more complex with styling, I recommend you look at another project for now. *However* it is my experience so far though that business users are more than willing to work with unstyled spreadsheets for reporting purposes and I also think that if you need professional grade looks, you should be delegating to a human _or_ use a professional, paid and supported library.

3. Example of use
-----------------

Things are really quite simple. Given a class Data like so:

    public class Data
    {
        public int ID { get; set; }
        public string Title { get; set; }
        public string AuthorName { get; set; }
    }

Then the below code:

    ...
    using(var fStream = File.Open("data.xlsx", FileMode.Create))
    {
        ExcelSaver.Save(dataToStore, fStream);
    }
    ...

will create an xlsx file with headers for the items passed in. Only public properties and fields are considered in this release.

4. Installation
---------------

To install it, use NuGet:
* From the Nuget package manager search for SuperSimple.Spreadsheets
* Alternativelly run Install-Package SuperSimple.Spreadsheets from the Package Manager Console

You can build it from source using Visual Studio. It has been tested to work with Visual Studio 2012, Visual Studio 2013 and Visual Studio 2015.
