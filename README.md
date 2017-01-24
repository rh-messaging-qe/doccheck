# DocCheck

This repository holds source code for OpenOffice/LibreOffice Calc macros used to download Red Hat documentation and create diffs. It simplifies performing documentation checks for our team (MSGQE). It may be also useful for a customer who is upgrading to a new minor release of a Red Hat product.

This repository is product-agnostic. There is another repo with MSGQE specific files.

## About this fork

This software holds a JVM rewrite of internal tool [jbossqe-eap-docs](http://git.app.eng.bos.redhat.com/git/jbossqe-eap-docs.git), a set of OOBasic macros used by the JBoss EAP QE team.

The reasons for the rewrite are as follows.

1. AFAICT, nobody on MSGQE team is comfortable with OOBasic.
2. LibreOffice Java API offers at least some code completion in IDEs.
3. Keeping the macros outside of the spreadsheet seemed better than embedding it.
4. Can have unittests for the code, so making changes will be faster.

### New features

* image links are replaced with embedded images
* for better performance, HTML is first mirrored with `wget`, then the mirror is imported

## About documentation testing workflow

This tool supports the following workflow.

1. Take a snapshot of what is published.
2. Create diff against previously checked snapshot.
3. Check all features described in modified sections.
4. Report issues to Jira.
5. Commit new snapshot to git.

Snapshots allow keeping track of what was checked and when. Otherwise the documentation would keep changing under our hands.

Diffs allow focusing our attention on new and modified instructions. Which is what needs checking the most.

Tracking the snapshots and diffs in git keeps record of what was checked and enables sharing it among the team.

## Supported platform

There is small amount of Linux-only code in the macro. This tool is developed and used for LibreOffice. It should work with OpenOffice as well.

## Running tests

The tests require running `soffice` that was given the `-accept` flag:

    soffice --headless --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"
    
(no window will open and nothing will be printed; this is headless execution)

Some tests open dialog boxes for visual inspection. These tests require soffice running without the `--headless` option.

## Installing the macros

This repository is built by Gradle. Use the wrapper (Unix or Windows version) to invoke Gradle. There is a custom task that installs the macro into your LibreOffice user directory.

    ./gradlew installMacro -Poo.scriptsdir=~/.config/libreofficedev/4/user/Scripts

### Tricks for running soffice

(from https://ask.libreoffice.org/en/question/2641/convert-to-command-line-parameter/)

It is possible to change user profile directory. This way, multiple soffice instances may run at the same time.

Do not forget to install the macros to all profiles where they are needed.

   soffice -env:UserInstallation=file:///tmp/tempprofile --headless --convert-to csv FileName 

## Seting up

1. Create a new Git repository.
2. Copy the file `sheet.fods` there and name it with the name of the product, for example `AMQ7.fods`. 
3. Edit the sheet and set correct URLs.

## Using the macros

The `.fods` document contains these sheets:

* Files -- downloading documentation, converting to FODT & PDF formats
* Comparisons -- detecting changed books, generating diff documents, exporting diffs to PDF format
* Settings -- configuration for the macros
* <PRODUCT_BUILD> -- new sheet is created after each download using Files sheet; it includes books downloaded with a particular build, specifies path to downloaded books

The following tasks can be performed:

### Download documentation for new EAP build - using 'Files' sheet:
1. In column 'B' specify URL which the book should be downloaded from (multiple URLs can be specified to download multiple books in one step)
2. Make sure column 'A' does not contain string 'Finished' (if it does, delete it)
3. Click 'Download documentation (click to download selected documentation)'
4. Enter version parameter (last used is pre-filled) and click 'OK'
5. Books specified in step 1 are downloaded and stored in EAP-<version> directory

### Compare newly downloaded books to previous build - using 'Comparisons' sheet:
1. If there is a row without 'Finished' in column 'A', then the book specified in the row has changed since last build
2. Click 'Click to compare documents which were not yet compared' to compare the changed books
3. Wait, it may take some time
4. Diff documents are stored in EAP-<version>/comparison directory

### Compare any two specified revisions - using 'Comparisons' sheet:
1. In column 'B:Original' specify path to ODT format of original book
2. In column 'C:New' specify path to ODT format of new book
3. In column 'D:Destination' specify path where the result should be stored; see bellow for diagram of directory structure
4. Click 'Click to compare documents which were not yet compared' to compare specified books
5. Diff documents are stored in directory specified in step 3

## Notes

* Macros must be enabled in the document; Default LibreOffice Security changes must be modified in order to do so.
* The LibreOffice GUI freezes while the macros run. There is probably nothing that can be done about it.

## Directory structure

Example of directory structure for EAP 6.3.0 (explanation below):

    jbossqe-eap-docs/
    |-- EAP6ComparisonSheet.ods
    |-- EAP7ComparisonSheet.ods
    |-- EAP-6.3.0.ER1
    |   `-- ODT and PDF formats of full books
    |-- EAP-6.3.0.ER2
    |   |-- comparison
    |   |   `-- ODT and PDF formats of comparison documents - compared to previous build (EAP-6.3.0.ER1)
    |   `-- ODF and PDF formats of full books
    `-- EAP-6.3.0.ER2-rebuild1
        |-- comparison
        |   |-- EAP-6.3.0.ER1
        |   |   `-- ODT and PDF formats of comparison documents - compared to specified build (EAP-6.3.0.ER1)
        |   `-- ODT and PDF formats of comparison documents - compared to previous build (EAP-6.3.0.ER2)
        `-- ODF and PDF formats of full books

* EAP6ComparisonSheet.ods -- tool for automated documents download and comparison, used for EAP 6
* EAP7ComparisonSheet.ods -- tool for automated documents download and comparison, used for EAP 7
* EAP-6.3.0.ER1 -- first build of particular EAP release, all books in PDF and ODT format are stored
* EAP-6.3.0.ER2 -- following builds (starting with ER2) in addition contain comparison directory
* comparison -- diff to previous version in PDF and ODT format; if it is needed, diff to another version can be generated and it is stored under comparison/<another_version> directory

Additional information
----------------------
Guidelines for EAP 6 documentation check in mojo: https://mojo.redhat.com/docs/DOC-953899
Guidelines for EAP 7 documentation check in mojo: https://mojo.redhat.com/docs/DOC-1049012


## Contacts

jdanek@redhat.com author of this Java fork

### Original authors

nziakova@redhat.com EAP documentation check coordinator
jkozana@redhat.com backup contact for EAP documentation check coordination