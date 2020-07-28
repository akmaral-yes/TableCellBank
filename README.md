# TableCellBank

## Introduction

TableCellBank is an image-based table strucure recognition dataset built with weak supervision from Word documents
on the internet, contains XXK high-quality labeled tables. 

The links to download Word documents from the Internet were provided by https://github.com/doc-analysis/TableBank

## Run the code
Download all the files from url_docx/url.csv:
```shell
$ python table_cell_from_docx/download_docx.py
```

Generate 100000 random colors with HEX and RGB codes:
```shell
$ python table_cell_from_docx/random_colors_generator.py --n 100000
```



## Structure Labels



## Implementation Details

The code was only tested on Windows 10 Pro
<!---## License --->

<!---## Citation --->

## References

- [Li et al., 2019] Li, M., L. Cui, S. Huang, F. Wei, M. Zhou, and Z. Li. Tablebank: Table benchmark
    for image-based table detection and recognition. arXiv preprint arXiv:1903.01949 .
