Python-oletools Changelog
=========================

- **2022-05-09 v0.60.1**:
    - olevba: 
      - fixed a bug when calling XLMMacroDeobfuscator (PR #737)
      - removed keyword "sample" causing false positives
    - oleid: fixed OleID init issue (issue #695, PR #696)
    - oleobj: 
      - added simple detection of CVE-2021-40444 initial stage
      - added detection for customUI onLoad
      - improved handling of incorrect filenames in OLE package (PR #451)
    - rtfobj: fixed code to find URLs in OLE2Link objects for Py3 (issue #692)
    - ftguess: 
      - added PowerPoint and XPS formats (PR #716)
      - fixed issue with XPS and malformed documents (issue #711)
      - added XLSB format (issue #758)
    - improved logging with common module log_helper (PR #449)
- **2021-06-02 v0.60**:
    - ftguess: new tool to identify file formats and containers (issue #680)
    - oleid: (issue #679)
        - each indicator now has a risk level
        - calls ftguess to identify file formats  
        - calls olevba+mraptor to detect and analyse VBA+XLM macros 
    - olevba:
        - when XLMMacroDeobfuscator is available, use it to extract and deobfuscate XLM macros
    - rtfobj:
        - use ftguess to identify file type of OLE Package (issue #682)
        - fixed bug in re_executable_extensions
    - crypto: added PowerPoint transparent password '/01Hannes Ruescher/01' (issue #627)
    - setup: XLMMacroDeobfuscator, xlrd2 and pyxlsb2 added as optional dependencies 
- **2021-05-07 v0.56.2**:
    - olevba:
        - updated plugin_biff to v0.0.22 to fix a bug (issues #647, #674)
    - olevba, mraptor:
        - added detection of Workbook_BeforeClose (issue #518)
    - rtfobj:
        - fixed bug when OLE package class name ends with null characters (issue #507, PR #648)
    - oleid:
        - fixed bug in check_excel (issue #584, PR #585)
    - clsid:
        - added several CLSIDs related to MS Office click-to-run issue CVE-2021-27058
        - added checks to ensure that all CLSIDs are uppercase (PR #678) 
- **2021-04-02 v0.56.1**:
    - olevba:
        - fixed bug when parsing some malformed files (issue #629)
    - oleobj:
        - fixed bug preventing detection of links 'externalReference', 'frame', 
          'hyperlink' (issue #641, PR #670)
    - setup:
        - avoid installing msoffcrypto-tool when platform is PyPy+Windows (issue #473)
        - PyPI version is now a wheel package to improve installation and avoid antivirus 
          false positives due to test files (issues #215, #398)
- **2020-09-28 v0.56**:
    - olevba/mraptor:
        - added detection of trigger _OnConnecting
    - olevba:
        - updated plugin_biff to v0.0.17 to improve Excel 4/XLM macros parsing
        - added simple analysis of Excel 4/XLM macros in XLSM files (PR #569)
        - added detection of template injection (PR #569)
        - added detection of many suspicious keywords (PR #591 and #569, see https://www.certego.net/en/news/advanced-vba-macros/)
        - improved MHT detection (PR #532)
        - added --no-xlm option to disable Excel 4/XLM macros parsing (PR #532)
        - fixed bug when decompressing raw chunks in VBA (issue #575)
        - fixed bug with email package due to monkeypatch for MHT parsing (issue #602, PR #604)
        - fixed option --relaxed (issue #596, PR #595)
        - enabled relaxed mode by default (issues #477, #593)
        - fixed detect_vba_macros to always return VBA code as
          unicode on Python 3 (issues  #455, #477, #587, #593)
        - replaced option --pcode by --show-pcode and --no-pcode,
          replaced optparse by argparse (PR #479)
    - oleform: improved form parsing (PR #532)
    - oleobj: "Ole10Native" is now case insensitive (issue #541)
    - clsid: added PDF (issue #552), Microsoft Word Picture (issue #571)
    - ppt_parser: fixed bug on Python 3 (issues #177, #607, PR #450)
- **2019-12-16 v0.55.2**:
    -  rtfobj:
        - removed "\rtf" from the list of destination control words (issue #522)
        - fixed process_file to detect Equation class (issue #525)
- **2019-12-03 v0.55**:
    - olevba:
        - added support for SLK files and XLM macro extraction from SLK
        - VBA Stomping detection
        - integrated pcodedmp to extract and disassemble P-code
        - detection of suspicious keywords and IOCs in P-code
        - new option --pcode to display P-code disassembly
        - improved detection of auto execution triggers
    - rtfobj: added URL carver for CVE-2017-0199
    - better handling of unicode for systems with locale that does not support UTF-8, e.g. LANG=C (PR #365)
    - tests: 
        - test files can now be encrypted, to avoid antivirus alerts (PR #217, issue #215)
        - tests that trigger antivirus alerts have been temporarily disabled (issue #215)
- **2019-05-22 v0.54.2**:
    - msoffcrypto-tool is now a required dependency (simplified install)
    - plugin_biff: fixed issues #428, #434 and #444, improved Python 3 support
    - olevba, msodde, crypto: improved handling of encrypted files (PR #441)
    - olevba: initialize VBA_Parser.xlm_macros (fixes #433)
    - various fixes (PR #446)
    - olevba and msodde now handle documents encrypted with common passwords such
      as 123, 1234, 4321, 12345, 123456, VelvetSweatShop automatically.
- **2019-04-09 v0.54.1**:
    - olevba: decompress_stream now accepts both bytes and bytearray (fixes #422)
- **2019-04-04 v0.54**:
    - olevba, msodde: added support for encrypted MS Office files 
    - olevba: added detection and extraction of XLM/XLF Excel 4 macros (thanks to plugin_biff from Didier Stevens' oledump)
    - olevba, mraptor: added detection of VBA running Excel 4 macros
    - olevba: detect and display special characters such as backspace
    - olevba: colorized output showing suspicious keywords in the VBA code
    - olevba, mraptor: full Python 3 compatibility, no separate olevba3/mraptor3 anymore
    - olevba: improved handling of code pages and unicode
    - olevba: fixed a false-positive in VBA macro detection
    - rtfobj: improved OLE Package handling, improved Equation object detection
    - oleobj: added detection of external links to objects in OpenXML
    - replaced third party packages by PyPI dependencies
- **2018-06-13 v0.53.1**:
    - rtfobj: fixed issue #316, whitespace after \bin on Python 3
    - olevba3: fixed #320, chr instead of unichr on python 3
    - olevba3: fixed #322, import reduce from functools
- **2018-05-30 v0.53**:
    - olevba and mraptor can now parse Word/PowerPoint 2007+ pure XML files (aka Flat OPC format)
    - improved support for VBA forms in olevba (oleform)
    - rtfobj now displays the CLSID of OLE objects, which is the best way to identify them. Known-bad CLSIDs such as MS Equation Editor are highlighted in red.
    - Updated rtfobj to handle obfuscated RTF samples.
    - rtfobj now handles the "\\'" obfuscation trick seen in recent samples such as https://twitter.com/buffaloverflow/status/989798880295444480, by emulating the MS Word bug described in https://securelist.com/disappearing-bytes/84017/
    - msodde: improved detection of DDE formulas in CSV files
    - oledir now displays the tree of storage/streams, along with CLSIDs and their meaning.
    - common.clsid contains the list of known CLSIDs, and their links to CVE vulnerabilities when relevant.
    - oleid now detects encrypted OpenXML files
    - fixed bugs in oleobj, rtfobj, oleid, olevba
- **2018-03-11 v0.52.2**:
    - Fixed issue #265 (error when installing on Python 3)
- **2018-02-18 v0.52**:
    - New tool [msodde](https://github.com/decalage2/oletools/wiki/msodde) to detect and extract DDE links from MS Office files, RTF and CSV;
    - Fixed bugs in olevba, rtfobj and olefile, to better handle malformed/obfuscated files;
    - Performance improvements in olevba and rtfobj;
    - VBA form parsing in olevba;
    - Office 2007+ support in oleobj.
- 2017-06-29 v0.51:
    - added the [oletools cheatsheet](https://github.com/decalage2/oletools/blob/master/cheatsheet/oletools_cheatsheet.pdf)
    - improved [rtfobj](https://github.com/decalage2/oletools/wiki/rtfobj) to handle malformed RTF files, detect vulnerability CVE-2017-0199
    - olevba: improved deobfuscation and Mac files support
    - [mraptor](https://github.com/decalage2/oletools/wiki/mraptor): added more ActiveX macro triggers
    - added [DocVarDump.vba](https://github.com/decalage2/oletools/blob/master/oletools/DocVarDump.vba) to dump document variables using Word
    - olemap: can now detect and extract [extra data at end of file](http://decalage.info/en/ole_extradata), improved display
    - oledir, olemeta, oletimes: added support for zip files and wildcards
    - many [bugfixes](https://github.com/decalage2/oletools/milestone/3?closed=1) in all the tools
    - improved Python 2+3 support
    
- 2016-11-01 v0.50: all oletools now support python 2 and 3.
    - olevba: several bugfixes and improvements.
    - mraptor: improved detection, added mraptor_milter for Sendmail/Postfix integration.
    - rtfobj: brand new RTF parser, obfuscation-aware, improved display, detect
    executable files in OLE Package objects.
    - setup: now creates handy command-line scripts to run oletools from any directory.
- 2016-06-10 v0.47: [olevba](https://github.com/decalage2/oletools/wiki/olevba) added PPT97 macros support,
improved handling of malformed/incomplete documents, improved error handling and JSON output,
now returns an exit code based on analysis results, new --relaxed option.
[rtfobj](https://github.com/decalage2/oletools/wiki/rtfobj): improved parsing to handle obfuscated RTF documents,
added -d option to set output dir. Moved repository and documentation to GitHub.
- 2016-04-19 v0.46: [olevba](https://github.com/decalage2/oletools/wiki/olevba)
does not deobfuscate VBA expressions by default (much faster), new option --deobf
to enable it. Fixed color display bug on Windows for several tools.
- 2016-04-12 v0.45: improved [rtfobj](https://github.com/decalage2/oletools/wiki/rtfobj)
to handle several [anti-analysis tricks](http://www.decalage.info/rtf_tricks),
improved [olevba](https://github.com/decalage2/oletools/wiki/olevba)
to export results in JSON format.
- 2016-03-11 v0.44: improved [olevba](https://github.com/decalage2/oletools/wiki/olevba)
to extract and analyse strings from VBA Forms.
- 2016-03-04 v0.43: added new tool [MacroRaptor](https://github.com/decalage2/oletools/wiki/mraptor) (mraptor)
to detect malicious macros, bugfix and slight improvements in [olevba](https://github.com/decalage2/oletools/wiki/olevba).
- 2016-02-07 v0.42: added two new tools oledir and olemap, better handling of malformed
files and several bugfixes in [olevba](https://github.com/decalage2/oletools/wiki/olevba),
improved display for [olemeta](https://github.com/decalage2/oletools/wiki/olemeta).
- 2015-09-22 v0.41: added new --reveal option to [olevba](https://github.com/decalage2/oletools/wiki/olevba),
to show the macro code with VBA strings deobfuscated.
- 2015-09-17 v0.40: Improved macro deobfuscation in [olevba](https://github.com/decalage2/oletools/wiki/olevba),
to decode Hex and Base64 within VBA expressions. Display printable deobfuscated strings by
default. Improved the VBA_Parser API. Improved performance.
Fixed [issue #23](https://github.com/decalage2/oletools/issues/23) with sys.stderr.
- 2015-06-19 v0.12: [olevba](https://github.com/decalage2/oletools/wiki/olevba) can now deobfuscate VBA
expressions with any combination of Chr, Asc, Val, StrReverse, Environ, +, &, using a VBA parser built with
[pyparsing](http://pyparsing.wikispaces.com). New options to display only the analysis results or only the macros source code.
The analysis is now done on all the VBA modules at once.
- 2015-05-29 v0.11: Improved parsing of MHTML and ActiveMime/MSO files in
[olevba](https://github.com/decalage2/oletools/wiki/olevba), added several suspicious keywords to VBA scanner
(thanks to @ozhermit and Davy Douhine for the suggestions)
- 2015-05-06 v0.10: [olevba](https://github.com/decalage2/oletools/wiki/olevba) now supports Word MHTML files
with macros, aka "Single File Web Page" (.mht) - see [issue #10](https://github.com/decalage2/oletools/issues/10) for more info
- 2015-03-23 v0.09: [olevba](https://github.com/decalage2/oletools/wiki/olevba) now supports Word 2003 XML files,
added anti-sandboxing/VM detection
- 2015-02-08 v0.08: [olevba](https://github.com/decalage2/oletools/wiki/olevba) can now decode strings
obfuscated with Hex/StrReverse/Base64/Dridex and extract IOCs. Added new triage mode, support for non-western
codepages with olefile 0.42, improved API and display, several bugfixes.
- 2015-01-05 v0.07: improved [olevba](https://github.com/decalage2/oletools/wiki/olevba) to detect suspicious
keywords and IOCs in VBA macros, can now scan several files and open password-protected zip archives, added a Python API,
upgraded OleFileIO_PL to olefile v0.41
- 2014-08-28 v0.06: added [olevba](https://github.com/decalage2/oletools/wiki/olevba), a new tool to extract VBA Macro
source code from MS Office documents (97-2003 and 2007+). Improved [documentation](https://github.com/decalage2/oletools/wiki)
- 2013-07-24 v0.05: added new tools [olemeta](https://github.com/decalage2/oletools/wiki/olemeta) and
[oletimes](https://github.com/decalage2/oletools/wiki/oletimes)
- 2013-04-18 v0.04: fixed bug in rtfobj, added documentation for [rtfobj](https://github.com/decalage2/oletools/wiki/rtfobj)
- 2012-11-09 v0.03: Improved [pyxswf](https://github.com/decalage2/oletools/wiki/pyxswf) to extract Flash objects from RTF
- 2012-10-29 v0.02: Added [oleid](https://github.com/decalage2/oletools/wiki/oleid)
- 2012-10-09 v0.01: Initial version of [olebrowse](https://github.com/decalage2/oletools/wiki/olebrowse) and pyxswf

See also the changelog in each source file for more details.
