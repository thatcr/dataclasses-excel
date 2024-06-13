Architecuture: we have multiple ways to discover a schema from an workbook
including a no-schema mode that maps named ranges as best it can.

There also multiple ways to extract the schema - openpyxl, xlwings, ms365
so the definitions have to be generic-ish and fall back.

The sheet itself can also contain type hints to control how ranges
are read, perhaps in the form of a json schema... or somthing more compact.

Named ranges of comment fields on them - use that to specify the type and maybe
field specified, so it forms the post-colon part

List[float] = Field(default=....)

Need some small DSL

