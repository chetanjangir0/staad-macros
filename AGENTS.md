- You have available documentiation for staad pro api in the open_staad_api_documentation folder
- before writing anything make sure that it is correct syntax and method according to the staad pro 2025 api and documentation

- The documentation folder doesn't include how to use api with python but here is the way to make it work
- also give me a conventinal commit style commit message for git after anything you work on

OS. Interpreting OpenSTAAD API Syntax for Python
The functions documented for the OpenSTAAD are given in the C++ syntax. It’s typically a simple process to interpret these for use in a Python program.

Classes and Methods
Classes syntax are listed in the format OSClassUI::Method. In Python, this should be instead given as Class.Method.

In the example, OpenSTAAD was instantiated using the variable "os". Classes were then called using the format os.Class (for example, the Geometry class was referenced using os.Geometry). In the C++ syntax used in the documentation, where you see "OSClassUI", instead simply use "Class" in Python.

The exception of course is the "root" OpenSTAAD functions, which will use whatever variable you have used to instantiate OpenSTAAD (looking to the example again, this would be os.GetBaseUnit). That is, the OpenSTAAD object acts as the class for these functions.

Note: In Python, you will have to use the _FlagAsMethod function to identify each OpenSTAAD method as a method once.
Return Values
The syntax listed indicates if a function has a return value with a leading "VARIANT". This simply means that function will return a value that typically will be stored into a variable. If the function syntax is listed with a leading "void", then there is no return value.

Note: A return value of a function is not to be confused with values stored in function parameters!
Variables and Parameters
The syntax listed typically indicates parameters as type "VARIANT FAR", which is not necessarily useful.

Instead, it is important that you read the parameter descriptions which will indicate what type of variable to use (i.e., "string", "long"). In Python, you don’t have to declare a variable type, it will be decided when the value is interpreted. Of course, it does matter if the variable contains a string or a number and what type of number when you want to use it in another function. OpenSTAAD may return integers or floating point decimal places.