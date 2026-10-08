# Excel Lookup Functions Practice

A structured collection of Excel lookup-function practice exercises
covering beginner, intermediate, and advanced lookup scenarios.

This repository documents hands-on practice with Excel formulas used to
retrieve, match, connect, and handle data across tables.

> **Note:** These exercises were practiced online. The repository
> therefore contains the reference material and learning documentation,
> but does not claim to contain an original `.xlsx` workbook.

## Topics Covered

### Beginner

-   VLOOKUP with exact matching
-   XLOOKUP
-   Lookup error handling
-   HLOOKUP
-   Basic lookup and reference concepts

### Intermediate

-   XLOOKUP with exact matching
-   XLOOKUP left lookup
-   XLOOKUP with default values
-   XLOOKUP wildcard search
-   VLOOKUP with IFNA
-   INDEX + MATCH
-   Two-way INDEX + MATCH
-   VLOOKUP approximate matching
-   Absolute references for reusable formulas

### Advanced

-   Two-way lookup with nested XLOOKUP
-   Returning multiple columns with XLOOKUP
-   Dynamic/spilled lookup results
-   Multi-dimensional lookup patterns

## Practice Exercises

  -----------------------------------------------------------------------
  \#                Exercise          Main Concept      Level
  ----------------- ----------------- ----------------- -----------------
  01                Find a customer's VLOOKUP           Beginner
                    city                                

  02                Look up a product XLOOKUP           Beginner

  03                Inventory stock   XLOOKUP exact     Beginner
                    lookup            match             

  04                Product SKU from  XLOOKUP left      Intermediate
                    name              lookup            

  05                Handle            XLOOKUP           Intermediate
                    discontinued SKUs `if_not_found`    

  06                Partial           XLOOKUP wildcard  Intermediate
                    product-name                        
                    search                              

  07                Missing-item      VLOOKUP + IFNA /  Beginner
                    handling          XLOOKUP           

  08                Manager lookup    INDEX + MATCH     Intermediate

  09                Shipping rate     INDEX + MATCH     Intermediate
                    matrix            two-way lookup    

  10                Product + quarter Nested XLOOKUP    Advanced
                    price                               

  11                Return Name,      XLOOKUP           Advanced
                    Email & Phone     multi-column      
                                      return            

  12                Product price     VLOOKUP           Beginner
                    lookup                              

  13                Employee          VLOOKUP           Beginner
                    department                          

  14                Quarterly sales   HLOOKUP           Beginner
                    target                              

  15                Tiered shipping   VLOOKUP           Intermediate
                    rate              approximate match 
  -----------------------------------------------------------------------

## Key Formulas Practiced

### VLOOKUP

``` excel
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

Used for retrieving a value from a table where the lookup field is in
the first column.

### XLOOKUP

``` excel
=XLOOKUP(lookup_value, lookup_array, return_array)
```

Used for flexible exact-match lookups and returning values from a
separate range.

### XLOOKUP with a fallback

``` excel
=XLOOKUP(F2,$A$2:$A$8,$B$2:$B$8,"Discontinued")
```

Used to return a meaningful result when a lookup value does not exist.

### Wildcard XLOOKUP

``` excel
=XLOOKUP("*"&B11&"*",$B$2:$B$9,$B$2:$B$9,,2)
```

Used to find a product containing a partial search term.

### INDEX + MATCH

``` excel
=INDEX($A$2:$A$9,MATCH(A13,$B$2:$B$9,0))
```

Used when the lookup value and return value are not arranged in a
VLOOKUP-friendly order.

### Two-way INDEX + MATCH

``` excel
=INDEX(B2:E6,MATCH(B8,A2:A6,0),MATCH(B9,B1:E1,0))
```

Used to find a value at the intersection of a selected row and column.

### Nested XLOOKUP

``` excel
=XLOOKUP(H2,A2:A8,XLOOKUP(H3,B1:E1,B2:E8))
```

Used to dynamically retrieve a value based on both a product and a
quarter.

### Multiple-column XLOOKUP

``` excel
=XLOOKUP(B3,A9:A15,B9:D15)
```

Used to return multiple related fields from a matching record.

### HLOOKUP

``` excel
=HLOOKUP(B4,A1:E2,2,FALSE)
```

Used when the lookup values are arranged horizontally.

### Approximate VLOOKUP

``` excel
=VLOOKUP(C5,$G$5:$H$9,2)
```

Used for tiered ranges such as shipping rates.

## Skills Demonstrated

-   Lookup and reference functions
-   Exact matching
-   Approximate matching
-   Left-side lookups
-   Two-way lookups
-   Wildcard searches
-   Lookup error handling
-   `IFNA`
-   Absolute cell references
-   Formula copying and reusable ranges
-   Dynamic array / spill behavior
-   Retrieving multiple fields from a record
-   Table-based data retrieval

## Learning Outcome

This practice helped build familiarity with selecting the appropriate
Excel lookup function for different data-retrieval scenarios.

The exercises progress from simple single-column lookups to more
advanced situations involving two-dimensional lookups, wildcard
searches, missing-value handling, and returning multiple fields.

## Repository Structure

``` text
excel-lookup-functions-practice/
│
├── README.md
│
├── Reference/
│   └── Excel Practice LOOKUP.pdf
│
└── Notes/
    └── Excel_Lookup_Topics.md
```

## Reference

The `Reference/` folder contains the PDF used for the practice exercises
and formula examples.

## Future Improvements

As I continue practicing Excel, this repository can be extended with:

-   Original `.xlsx` practice workbooks
-   Screenshots of completed exercises
-   Excel dashboards
-   Pivot Tables and Pivot Charts
-   Data-cleaning exercises
-   Business-focused Excel analysis projects
-   Power Query practice

## Profile Relevance

These Excel skills complement data-analysis work involving data
retrieval, validation, cleaning, reporting, and exploratory analysis.

**Focus:** Excel • Data Analysis • Lookup Functions • Data Retrieval •
Reporting

