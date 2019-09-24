# Word-to-LaTeX

This repository contains a guide on how to convert Word documents into $\LaTeX$ and Markdown.

The repository comprises a number of files and folders.

1. The base directory
  - 'word-to-latex.Rmd' -- an R markdown file that, when knit to html, produces a presentation. Produced with the R package 'xaringan.'
  - 'word-to-latex.html' -- the html presentation of above
  - 'sfah.css' -- a css style that is used within the Rmd presentation file above.
  
2. '/written guides'
  - 'Guide for Converting Documents from Word to Latex.docx' -- a Word document version of the presentation; somewhat more thorough
  - 'steps to take to convert word to latex.txt' -- a quick and dirty big picture view of the steps involved, focusing on manipulation post-conversion into a $\LaTeX$ file.
  
3. '/VBA scripts'
  - 'Word-to-Latex-VBA.txt' -- text file containing three VBA macros
    - one converts citations to LaTeX format, another to Markdown format
    - the last finds figure cross-references and changes these to LaTeX code with user-specified labels
    
4. '/csl styles'
  - 'Word-to-Latex.csl' -- the CSL style necessary to show Zotero citations as LaTeX format within Word docs
  - 'Word-to-Markdown.csl' -- the CSL style for Markdown
  
5. '/libs'
  - files necessary to produce the 'word-to-latex' presentation (e.g., remark.js scripts, on which 'xaringan' depends)
  
6. '/figures'
  - three pngs that are used in the 'word-to-latex' presentation