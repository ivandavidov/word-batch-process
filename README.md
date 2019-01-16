# MS Word Placeholder Processor

### Prerequisites

* [Apache POI](https://poi.apache.org/)
* [Apache Commons IO](https://commons.apache.org/proper/commons-io)

A helper program which automatically processes placeholders in MS Word files.

### Compile

1. Create new folder ``lib`` in the main project.
2. Download [Apache POI](https://poi.apache.org/) and extract all libraries in ``lib``.
3. Download [Apache Commons IO](https://commons.apache.org/proper/commons-io) and extract all libraries in ``lib``.
4. Edit ``compile.cmd`` and set the appropriate value for ``JAVA_HOME`` (must be JDK).
5. Run ``compile.cmd``.

### How to use

1. Create MS Excel file with arbitrary name (e.g. 'placeholders.xlsx').
2. Create table in the first sheet. The first row contains the placeholders. Each next row contains the placeholder values.

   Example:

   ```
   id | name   | age
    1 | John   |  25
    2 | David  |  31
    3 | Marina |  20
   ```

3. Create MS Word document with arbitrary name. This will be your template. Surround each placeholder with ``__``. Don't use speciall formatting/highlighting around the placeholders. The template file itself can contain one or more placeholders.

   
   Sample file name: ``template__id__test.docx``
   
   Sample template text:
   
   ```
   __name__ is __age__ years old with ID __id__.
   ```
   
4. Edit ``wordbatch.cmd`` and set the appropriate value for ``JAVA_HOME`` (either JDK or JRE).
5. Run ``wordbatch.cmd``.
6. The generated MS Word files are in the newly created folder ``result``.
