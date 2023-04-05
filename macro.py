import aspose.words as aw

# Read payload of agents
src=''
f= open("./agents/W_macro.vba", "r")
a =True
# src=f.read()

# Load Word document.
doc = aw.Document("document.docm")

# Create VBA project
project = aw.vba.VbaProject()
project.name = "DropperProj"
doc.vba_project = project

# Create a new module and specify a macro source code.
module = aw.vba.VbaModule()
module.name = "DropperModule"
module.type = aw.vba.VbaModuleType.PROCEDURAL_MODULE
# module.source_code=src
module.source_code='Test crazy macro'
# while a:
# 	line = f.readline(1)
# 	if line != '':
# 		module.source_code+=line
# 	else:
# 		a =False
# f.close()
# print(module.source_code)

# Add module to the VBA project.
doc.vba_project.modules.add(module)

# Save document.
doc.save("readme.docm")
