msodde
======

msodde is a script to parse MS Office documents (e.g. Word, Excel, RTF, XML), to detect and extract **DDE links** such as 
**DDEAUTO**, that have been used to run malicious commands to deliver malware.
It also supports CSV files, which may contain Excel formulas to run executable files using DDE (technique known as "CSV injection").
For Word documents, it can extract all the other fields, and identify suspicious ones.

Supported formats:
- Word 97-2003 (.doc, .dot), Word 2007+ (.docx, .dotx, .docm, .dotm)
- Excel 97-2003 (.xls), Excel 2007+ (.xlsx, .xlsm, .xlsb)
- RTF
- CSV (exported from / imported into Excel)
- XML (exported from Word 2003, Word 2007+, Excel 2003, Excel 2007+)

For Word documents, msodde detects the use of QUOTE to obfuscate DDE commands (see 
[this article](http://staaldraad.github.io/2017/10/23/msword-field-codes/)), and deobfuscates
it automatically. 

Special thanks to Christian Herdtweck and Etienne Stalmans, who contributed large parts of 
the code.

msodde can be used either as a command-line tool, or as a python module
from your own applications.

It is part of the [python-oletools](http://www.decalage.info/python/oletools) package.

## References about DDE exploitation

- https://www.contextis.com/blog/comma-separated-vulnerabilities
- http://www.exploresecurity.com/from-csv-to-cmd-to-qwerty/
- https://pwndizzle.blogspot.nl/2017/03/office-document-macros-ole-actions-dde.html
- https://sensepost.com/blog/2017/macro-less-code-exec-in-msword/
- http://staaldraad.github.io/2017/10/23/msword-field-codes/
- https://xorl.wordpress.com/2017/12/11/microsoft-excel-csv-code-execution-injection-method/
- http://georgemauer.net/2017/10/07/csv-injection.html
- http://blog.7elements.co.uk/2013/01/cell-injection.html
- https://appsecconsulting.com/blog/csv-formula-injection
- https://www.owasp.org/index.php/CSV_Injection

## Usage

```text
usage: msodde.py [-h] [-j] [--nounquote] [-l LOGLEVEL] [-p PASSWORD] [-d] [-f]
                 [-a]
                 FILE

positional arguments:
  FILE                  path of the file to be analyzed

optional arguments:
  -h, --help            show this help message and exit
  -j, --json            Output in json format. Do not use with -ldebug
  --nounquote           don't unquote values
  -l LOGLEVEL, --loglevel LOGLEVEL
                        logging level debug/info/warning/error/critical
                        (default=warning)
  -p PASSWORD, --password PASSWORD
                        if encrypted office files are encountered, try
                        decryption with this password. May be repeated.

Filter which OpenXML field commands are returned:
  Only applies to OpenXML (e.g. docx) and rtf, not to OLE (e.g. .doc). These
  options are mutually exclusive, last option found on command line
  overwrites earlier ones.

  -d, --dde-only        Return only DDE and DDEAUTO fields
  -f, --filter          Return all fields except harmless ones
  -a, --all-fields      Return all fields, irrespective of their contents
```

**New in v0.54:** the -p option can now be used to decrypt encrypted documents using the provided password(s).

### Examples

Scan a single file:

```text
msodde file.doc
```

Scan a Word document, extracting *all* fields:

```text
msodde -a file.doc
```


--------------------------------------------------------------------------
    
## How to use msodde in Python applications

This is work in progress. The API is expected to change in future versions. 


--------------------------------------------------------------------------

python-oletools documentation
-----------------------------

- [[Home]]
- [[License]]
- [[Install]]
- [[Contribute]], Suggest Improvements or Report Issues
- Tools:
	- [[mraptor]]
	- [[msodde]]
	- [[olebrowse]]
	- [[oledir]]
	- [[oleid]]
	- [[olemap]]
	- [[olemeta]]
	- [[oleobj]]
	- [[oletimes]]
	- [[olevba]]
	- [[pyxswf]]
	- [[rtfobj]]
