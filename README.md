# excel_filter

#### help
.\excel_filter.exe --help
```
excel_filter 0.1.0
None
Excel File Filter

USAGE:
    excel_filter.exe -s <Source Path>... -k <Source Sheet Name>... -t <Target Path>... -l <Target Sheet Name>... -c <Matching Column>... -m <Matching String>...

OPTIONS:
    -c <Matching Column>...          Input matching column, starting from 0.
    -h, --help                       Print help information
    -k <Source Sheet Name>...        Input source sheet name.
    -l <Target Sheet Name>...        Input target sheet name.
    -m <Matching String>...          Input matching string.
    -s <Source Path>...              Input source path.
    -t <Target Path>...              Input target path.
    -V, --version                    Print version information
```
#### run
.\excel_filter.exe -s .\file.xlsx -k "Sheet1" -t .\newfile.xlsx -c 1 -m "Changes done successfully" -l "filtered"
