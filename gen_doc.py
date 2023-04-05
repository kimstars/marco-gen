import aspose.words as aw

# create document object
doc = aw.Document()

# create a document builder object
builder = aw.DocumentBuilder(doc)

# set list formatting
builder.list_format.apply_number_default()

# insert item
builder.writeln("Item 1")
builder.writeln("Item 2")

# set indentation for next level of list
builder.list_format.list_indent()
builder.writeln("Item 2.1")
builder.writeln("Item 2.2")

# indent again for next level
builder.list_format.list_indent()
builder.writeln("Item 2.2.1")
builder.writeln("Item 2.2.2")

# outdent to get back to previous level
builder.list_format.list_outdent()
builder.writeln("Item 2.3")

# outdent again to get back to first level
builder.list_format.list_outdent()
builder.writeln("Item 3")

# remove numbers
builder.list_format.remove_numbers()

# save document
doc.save("document.docm")