# File formats, Techniques and Tools

This table shows the various techniques that can be used in malicious
documents to trigger code execution, and the file formats in which they
can be embedded. The last row suggests tools that can detect and analyse
each technique.

Each technique is described below the table.

This is work in progress, not all combinations have been thoroughly
tested.

<table>
<thead>
<tr class="header">
<th><strong>File Format / Technique</strong></th>
<th><strong>VBA Macros</strong></th>
<th><strong>Excel 4 / XLM Macros</strong></th>
<th><strong>DDE</strong></th>
<th><strong>OLE Objects</strong></th>
<th><strong>Package OLE Objects</strong></th>
<th><p><strong>Remote Template</strong></p>
<p>(<a href="https://attack.mitre.org/techniques/T1221/">T1221</a>)</p></th>
<th><strong>Remote OLE object</strong></th>
<th><strong>customUI (remote macro)</strong></th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td>Word 97-2003 (DOC)</td>
<td>X</td>
<td>-</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>?</td>
</tr>
<tr class="even">
<td>Word 2007+ (DOCX)</td>
<td>-</td>
<td>-</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>X</td>
</tr>
<tr class="odd">
<td>Word 2007+ macro-enabled (DOCM)</td>
<td>X</td>
<td>-</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>X</td>
</tr>
<tr class="even">
<td>Excel 97-2003 (XLS)</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>?</td>
<td>X</td>
<td>?</td>
</tr>
<tr class="odd">
<td>Excel 2007+ (XLSX)</td>
<td>-</td>
<td>?</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>?</td>
<td>X</td>
<td>X</td>
</tr>
<tr class="even">
<td>Excel 2007+ macro-enabled (XLSM)</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>?</td>
<td>X</td>
<td>X</td>
</tr>
<tr class="odd">
<td><p>Excel 2007+ Binary</p>
<p>(XLSB)</p></td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>?</td>
<td>X</td>
<td>X</td>
</tr>
<tr class="even">
<td>PowerPoint 97-2003 (PPT)</td>
<td>X</td>
<td>-</td>
<td>?</td>
<td>X</td>
<td>X</td>
<td>?</td>
<td>X</td>
<td>?</td>
</tr>
<tr class="odd">
<td>PowerPoint 2007+ (PPTX)</td>
<td>-</td>
<td>-</td>
<td>?</td>
<td>X</td>
<td>X</td>
<td>?</td>
<td>X</td>
<td>X</td>
</tr>
<tr class="even">
<td>PowerPoint 2007+ macro-enabled (PPTM)</td>
<td>X</td>
<td>-</td>
<td>?</td>
<td>X</td>
<td>X</td>
<td>?</td>
<td>X</td>
<td>X</td>
</tr>
<tr class="odd">
<td>RTF</td>
<td>-</td>
<td>-</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>?</td>
</tr>
<tr class="even">
<td>CSV</td>
<td>-</td>
<td>-</td>
<td>X</td>
<td>-</td>
<td>-</td>
<td>-</td>
<td>-</td>
<td>-</td>
</tr>
<tr class="odd">
<td>SLK</td>
<td>-</td>
<td>X</td>
<td>X</td>
<td>-</td>
<td>-</td>
<td>-</td>
<td>-</td>
<td>-</td>
</tr>
<tr class="even">
<td>MHT (from Word)</td>
<td>X</td>
<td>?</td>
<td>?</td>
<td>X</td>
<td>X</td>
<td>?</td>
<td>?</td>
<td>?</td>
</tr>
<tr class="odd">
<td>MHT (from Excel)</td>
<td>?</td>
<td>?</td>
<td>?</td>
<td>?</td>
<td>?</td>
<td>?</td>
<td>?</td>
<td>?</td>
</tr>
<tr class="even">
<td>Word 2003 XML</td>
<td>X</td>
<td>-</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>?</td>
<td>?</td>
<td>?</td>
</tr>
<tr class="odd">
<td>Word 2016 XML</td>
<td>X</td>
<td>-</td>
<td>X</td>
<td>X</td>
<td>X</td>
<td>?</td>
<td>?</td>
<td>?</td>
</tr>
<tr class="even">
<td>Excel 2003 XML</td>
<td>?</td>
<td>?</td>
<td>?</td>
<td>?</td>
<td>?</td>
<td>?</td>
<td>?</td>
<td>?</td>
</tr>
<tr class="odd">
<td>Publisher (PUB)</td>
<td>X</td>
<td>-</td>
<td>?</td>
<td>X</td>
<td>X</td>
<td>?</td>
<td>?</td>
<td>?</td>
</tr>
<tr class="even">
<td>Visio (VSDX)</td>
<td>X</td>
<td>-</td>
<td>?</td>
<td>?</td>
<td>?</td>
<td>?</td>
<td>?</td>
<td>?</td>
</tr>
<tr class="odd">
<td><strong>Tools</strong></td>
<td><p><a href="https://github.com/decalage2/oletools/wiki/olevba">olevba</a></p>
<p><a href="https://github.com/decalage2/oletools/wiki/mraptor">mraptor</a></p>
<p><a href="https://github.com/decalage2/ViperMonkey">ViperMonkey</a></p>
<p><a href="https://blog.didierstevens.com/programs/oledump-py/">oledump</a></p></td>
<td><p><a href="https://github.com/decalage2/oletools/wiki/olevba">olevba</a></p>
<p><a href="https://blog.didierstevens.com/programs/oledump-py/">oledump</a></p>
<p><a href="https://github.com/DissectMalware/XLMMacroDeobfuscator">XLMMacro Deobfuscator</a></p></td>
<td><a href="https://github.com/decalage2/oletools/wiki/msodde">msodde</a></td>
<td><p><a href="https://github.com/decalage2/oletools/wiki/oleobj">oleobj</a></p>
<p><a href="https://github.com/decalage2/oletools/wiki/rtfobj">rtfobj</a></p></td>
<td><p><a href="https://github.com/decalage2/oletools/wiki/oleobj">oleobj</a></p>
<p><a href="https://github.com/decalage2/oletools/wiki/rtfobj">rtfobj</a></p></td>
<td><a href="https://github.com/decalage2/oletools/wiki/oleobj">oleobj</a></td>
<td><a href="https://github.com/decalage2/oletools/wiki/oleobj">oleobj</a></td>
<td><a href="https://github.com/decalage2/oletools/wiki/oleobj">oleobj</a></td>
</tr>
</tbody>
</table>

## Techniques

### VBA Macros

VBA (Visual Basic for Applications) is a programming language used to
automate tasks in Microsoft Office applications since 1997. VBA macros
may be embedded into Word documents, Excel spreadsheets, PowerPoint
presentations, etc. A VBA macro can be triggered automatically when
opening or closing a file (after clicking “Enable Content”), and it can
execute any action on the system such as dropping a file, executing a
command, calling any DLL or ActiveX object. In practice, a VBA macro is
just as powerful as any EXE.

More info: <https://decalage.info/en/bheu2019>

In 2022, Microsoft plans to disable VBA macros in files coming from the
Internet, starting with Office365:
<https://techcommunity.microsoft.com/t5/microsoft-365-blog/helping-users-stay-safe-blocking-internet-macros-by-default-in/ba-p/3071805>

### Excel 4 / XLM Macros

Excel 4 Macros offer similar functionality and risks as VBA macros, but
the language and the engine are completely different. XLM Macros are
composed of formulas in cells, and they only run on Excel.

Some
    references:

  - <https://outflank.nl/blog/2018/10/06/old-school-evil-excel-4-0-macros-xlm/>

  - <https://outflank.nl/blog/2019/10/30/abusing-the-sylk-file-format/>

  - <https://www.lastline.com/labsblog/evolution-of-excel-4-0-macro-weaponization/>

XLM Macros are disabled by default since July 2021:
<https://techcommunity.microsoft.com/t5/excel-blog/excel-4-0-xlm-macros-now-restricted-by-default-for-customer/ba-p/3057905>

### DDE

DDE (Dynamic Data Exchange) is a Microsoft protocol to enable data
sharing between applications. In some applications such as Word and
Excel, it has been found that it was possible to abuse DDE to launch any
command. It is even possible to trigger code execution in Excel from a
simple CSV file, by embedding specific formulas.

Some
    references:

  - <https://www.contextis.com/us/blog/comma-separated-vulnerabilities>

  - <https://sensepost.com/blog/2016/powershell-c-sharp-and-dde-the-power-within/>

  - <http://www.exploresecurity.com/from-csv-to-cmd-to-qwerty/>

  - <https://sensepost.com/blog/2017/macro-less-code-exec-in-msword/>

  - <https://docs.microsoft.com/en-us/security-updates/securityadvisories/2017/4053440>

The ability to launch arbitrary commands using DDE has been
progressively disabled by default in Word (2017) and then Excel (2022):
<https://msrc.microsoft.com/update-guide/en-US/vulnerability/ADV170021>

### OLE Objects

OLE is a Microsoft protocol used to embed data from one application into
a file from another application. For example, it can be used to embed an
Excel chart into a Word document. In general, OLE objects cannot trigger
the execution of arbitrary code or commands. However, in the past many
vulnerabilities have been exploited thanks to OLE objects. For example,
the vulnerability
[CVE-2017-11882](https://cve.mitre.org/cgi-bin/cvename.cgi?name=CVE-2017-11882)
in the MS Equation Editor has been actively exploited by embedding
malformed Equation OLE objects into Word and RTF documents.

### OLE Package objects

TODO

### Remote Template

TODO (T1221)

### Remote OLE object 

TODO

### customUI (remote macro)

TODO
