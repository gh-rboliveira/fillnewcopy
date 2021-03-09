# FillNewCopy

FillNewCopy is a GUI tool that will generate a new copy of a templated filled with the information introduced.  

FillNewCopy is based in:

* Configuration file (A Json file)
* Template File (A Word document as template)

It will generate the GUI layout based in the config file.

Also based on the configuration file it will generate a new copy of the template file, replacing the placeholders (**<KEY_IN_UPPERCASE>**) defined in the configuration as **key** for each field.

**IMPORTANT NOTE**: The placeholder should have the same style. if the placeholder holds different styling, it **WILL NOT** work

Example

----

 

For the below configuration:

```json
{
   "app_name":"Generate",
   "template_file":"template.docx",
   "filename":{
      "type":"field",
      "value":"client_number"
   },
   "file_prefix":"Client_",
   "fields":{
      "person_name":{
         "label":"Name",
         "type":"str",
         "default":"",
         "required":1
      },
      "client_number":{
         "label":"Client Number",
         "type":"str",
         "default":"",
         "required":1
      },
      "age":{
         "label":"Age",
         "type":"list",
         "options":[
            "<18",
            "18-65",
            "65+"
         ],
         "default":"18-65",
         "required":1
      }
   }
}
```

It will generate the following window:

![Window]("/print_screens/Layout.JPG")

In our document we must have the placeholders "**<PERSON_NAME>**","**<CLIENT_NUMBER>**","**<AGE>**"

![Before](https://ibb.co/RhsJvm9)

After clicking **Build** a new file named Client_10001.docx will be created and filled with the data provided

![After](https://ibb.co/YR0cn25)



## Configuration File

 

### GENERAL OPTIONS:

 

At the ***root*** of your json the possible keys are:

 

```json
{
   "app_name":"Generate",
   "template_file":"file_template.docx",
   "filename":{
      "type":"field",
      "value":"client_number"
   },
   "file_prefix":"",
   "file_posfix":"",
   "date_format":"%d-%m-%Y",
   "fields":{
      ...
   }
}
```



field  | Mandatory | What does it do
------------- | ------------- | -------------
app_name  | NO | As the name suggests, it replaces the name of this App
template_file | YES | The name of the template (Word Document) to be used, the file should be placed in the root of the project
filename | NO | You can configure if you want your filename to be generated based in a **field** or use a **static** filename by indicating the type. For both a **value** should be provided
file_prefix | NO | prefix for the generated filename
file_posfix | NO | posfix for the generated filename
date_format | NO | When using date fields you can specify a date format ([Python date formats](https://docs.python.org/3/library/datetime.html#strftime-and-strptime-format-codes))
fields | YES | The definition of the fields, see below

 

### FIELDS OPTIONS

 

The **key** of each field is considered the field identifier, inside of each **field** the possible keys are:

 

```json

{
   ...
   "fields":{
      "person_name":{
         "label":"Name",
         "type":"str",
         "default":"",
         "required":1
      },
      "age":{
         "label":"Age",
         "type":"list",
         "options":[
            "<18",
            "18-65",
            "65+"
         ],
         "default":"18-65",
         "required":1
      }
   }
}

```

 

Where the **key** (example: "person_name") is considered the **field identifier**
 

field  | Mandatory | What does it do
------------- | ------------- | -------------
label | NO |To be used in the GUI. The default will be the key in uppercase
type | NO | You can define 3 types of field **str** (String), **date** (Date), **list** (A list of values to be picked up). Undefined or non provided types will be interpreted as Strings (**str**)
default| NO | A default value for the field
required | NO | If a field should be required in the GUI
options | NO | **(Only for list type)** an array of options for the list

## Usage

```bash
λ python fillnewcp.py
```

## Contributing

This tool was made more for learning purposes. Please feel free to fork, or even contribute.
 

## License

[MIT](https://choosealicense.com/licenses/mit/)