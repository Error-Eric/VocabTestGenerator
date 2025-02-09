### VocabTest Generator (v0.0.1)

This is a simple project (customized for some language exam prep school) that generates vocabulary test based on the vocabulary list given. Make sure that your vocabulary list is given in the following format (you can see sample.xlsx):

- An readable excel file.

- There's no units, or each unit begins with a cell ``unit x``, coming above the content of the unit.

- Every cell consists of only letters are the words and the meaning is exactly one cell to the right.

- To avoid bugs, make sure no cells are merged.

I didn't add parameter check in some places (since it's only a simple project and I have no idea what weird input might be given.) Just remember to check parameters yourself.

Download ``dist.zip`` and unpack it. Hopefully it's easy enough to use.

--------------

This project is based on 

```plain
Package    Version
---------- -------
et_xmlfile 2.0.0
openpyxl   3.1.5
pip        24.0
tk         0.1.0
```

Few people are going to use this project. It's highly unlikely that I won't keep updating this.

#### Update log

Feb 9 2025. Better Saving features and printable xlsx file.