# TableCellBank

## Introduction

TableCellBank is table strucure recognition dataset which contains 78553 table images and corresponding annotations obtained from Word source files.
<!---contains XXK high-quality labeled tables. --->

The links to download Word documents from the Internet were provided by https://github.com/doc-analysis/TableBank

We provide only a code to build TableCellBank from these Word documents.

## Run the code
Download all the files from url_docx/url.csv:
```shell
$ python table_cell_from_docx/download_docx.py
```

Generate 100000 random colors with HEX and RGB codes:
```shell
$ python table_cell_from_docx/random_colors_generator.py --n 100000
```

Retrieve table images from downloaded Word documents and build corresponding ground truth:
```shell
$ python table_cell_from_docx/table_cell_from_docx.py --do run --multiproc
```

## Structure Labels
Every document name is a randomly generated uuid.
To build a table name a document name and a page number that contains this table and the order number of this table on the page are concatenated with an underscore.

The ground truth files are created per document.
These files contain table names, their locations (x1,y1,x2,y2), cell positions, horizontal and vertical line positions.

```json
{
    "00bd111b-a411-4176-91f7-957bd80cf6fe_48_0.png": {
        "cells": [
            [
                5,
                5,
                438,
                114
            ],
            [
                448,
                5,
                414,
                114
            ],
            ...
        ],
        "horizontal_lines": [
            [
                5,
                5,
                862,
                5
            ],
            ...
        ],
        "loc": [
            807,
            529,
            867,
            371
        ],
        "vertical_lines": [
            [
                5,
                5,
                5,
                366
            ],
            ...
        ]
    },
  ...
}
```

## Implementation Details

The code was only tested on Windows 10 Pro
<!---## License --->

<!---## Citation --->

## References

- [Li et al., 2019] Li, M., L. Cui, S. Huang, F. Wei, M. Zhou, and Z. Li. Tablebank: Table benchmark
    for image-based table detection and recognition. arXiv preprint arXiv:1903.01949 .
