officeXtract
==============

Python - Office Open XML file string extraction tool

EXAMPLE USAGE;

EXAMPLE USAGE;

`
$ python officextract.py [summary] filename.xlsx`

    Extracts all unique strings from Office .x files and prints them to stdout'

    Optional argument summary prints out only information about the file\n'
`


Will output either all strings found in the document or a summary of what was found - e.g.;

`
Summary for /home/somefolder/somefile.docx

  Processed files;
    word/document.xml
    word/footer1.xml
    word/footnotes.xml
    word/endnotes.xml
    word/theme/theme1.xml
    word/charts/chart1.xml
    word/settings.xml
    word/styles.xml
    word/numbering.xml
    customXml/itemProps1.xml
    customXml/item1.xml
    docProps/core.xml
    word/fontTable.xml
    word/webSettings.xml
    word/stylesWithEffects.xml
    docProps/app.xml

  Ignored files;
    [Content_Types].xml
    _rels/.rels
    word/_rels/document.xml.rels
    word/media/image10.png
    word/media/image6.png
    word/media/image7.png
    word/media/image8.png
    word/media/image9.gif
    word/media/image13.jpg
    word/media/image12.gif
    word/media/image11.jpeg
    word/media/image4.png
    word/media/image3.png
    word/charts/_rels/chart1.xml.rels
    word/media/image5.png
    word/media/image2.png
    word/media/image1.png
    customXml/_rels/item1.xml.rels

  Phrases/lines: 1008
  Single words with ignored characters at beginning: 8
  Blank lines: 10558
  Words shorter than 3 chars: 408
`
