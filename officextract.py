#!/usr/bin/env python
"""Office Open XML file string extractor.  Nick Besant 2014-2015 hwf@fesk.net

EXAMPLE USAGE;
$ python officextract.py [summary] filename.xlsx
    Extracts all unique strings from Office .x files and prints them to stdout'
    Optional argument summary prints out only information about the file\n'

"""

import sys,zipfile
from xml.etree.cElementTree import XML

if len(sys.argv)==1:
    print 'python officextract.py [summary] filename.xlsx'
    print '   Extracts all unique strings from Office .x files and prints them to stdout'
    print '   Optional argument summary prints out only information about the file\n'
    sys.exit()

summary=False

if sys.argv[1]=='summary':
    infile=sys.argv[2]
    summary=True
else:
    infile=sys.argv[1]


# Very basic error checking
try:
    sourcefile=zipfile.ZipFile(infile)
except Exception as e:
    print 'Extract failed, is it an Office Open XML file?\n\n\n{0}'.format(e)
    sys.exit()

outset=set()

fnames=[]
fignores=[]
fcount=0
wignores=0
wshort=0
wblanks=0

# iterate through the files within the zip
for fname in sourcefile.namelist():
    fcount+=1
    # does this look like a useful file
    if fname[-3:]=='xml' and fname[0]!='[':
        fnames.append(fname)
        # try parsing the file
        try:
            xmltree=XML(sourcefile.read(fname))
        except Exception as e:
            print u'Error parsing file {0} in document {1}: {2}'.format(fname,infile, e)
            sys.exit()
        for para in xmltree.iter():
            if para.text:
                outtext=para.text.strip()
                # only bother with strings longer than 3 chars and not beginning with _, $ or ?
                if len(outtext)>3:
                    if ' ' not in outtext and outtext[0] in "'_$?":
                        wignores+=1
                    else:
                        outset.add(outtext)
                else:
                    wshort+=1
            else:
                wblanks+=1
    else:
        fignores.append(fname)

sourcefile.close()

if summary:
    print '\nSummary for {0}\n'.format(infile)
    print '  Processed files;\n    {0}\n'.format('\n    '.join(x for x in fnames))
    print '  Ignored files;\n    {0}\n'.format('\n    '.join(x for x in fignores))
    print '  Phrases/lines: {0}'.format(len(outset))
    print '  Single words with ignored characters at beginning: {0}'.format(wignores)
    print '  Blank lines: {0}'.format(wblanks)
    print '  Words shorter than 3 chars: {0}\n\n'.format(wshort)
else:
    print u'\n'.join(outset).encode('utf-8')