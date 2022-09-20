# WikiToDoc

Export your Azure DevOps Wiki to a series of word documents. WikiToDoc converts markdown files to docx format Microsoft Word documents. Each page is converted into a separate word document.

## Usage
Azure DevOps lets you checkout your wiki pages as a git repo. Clone that to a local directory.

WikiToDoc.exe takes one parameter, the path of your local wiki directory. `WikiToDoc.exe C:\Code\MyProjectWiki`

## Prerequisites

Install [pandoc](https://pandoc.org/) and make sure its in your path.
