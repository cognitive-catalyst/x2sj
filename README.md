# x2sj (Excel 2 SOLR JSON transform utility)
Node utility to transform Excel spreadsheets into JSON for consumption into the Retrieve (SOLR) and Rank service on Bluemix.

x2sj is a component utility of the Watson Catalyst program.


### Prerequisites

You need to have node installed. Install it from here: http://nodejs.org


### How to install

Install the package globally using npm. Note, on Linux or Mac you might need to include 'sudo' before the command

```
npm install x2sj -g
```


### Usage

As the utility is stored globally, you may execute it like any other cli command:

```
% x2sj
Usage:
  x2sj [OPTIONS] [ARGS]

Options:
  -i, --file FILE        An XLSX file to process
  -o, --output FILE      Output file in JSON format
  -h, --help             Display help and usage details
```

### Data

The utility operates on Excel Spreadsheets. The first column must be the operation to execute ('add', 'remove', etc). The rest of the column headers are taken to be valid attributes that will be added to the document in SOLR. If your SOLR schema does not contain a corresponding entry, then your JSON will fail to load:(

Example:

| operation  |  id  |  curratedon  |  title             |
|------------|------|--------------|--------------------|
| add        | 1    | 20151030     | doc1 title part 1  |
| add        | 1    | 20151030     | doc1 title part 2  |
| remove     | 2    | 20141215     | doc100             |

### People

The original author is [Chris Madison](https://github/tankcdr).

### License

Apache-2.0
