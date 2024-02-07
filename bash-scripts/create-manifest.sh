#!/bin/bash


# Define the XML file name
original_xml_file="manifest-template.xml"

# Define the output XML file name
modified_xml_file="$1"

# Define the replacement values
appUri="$2"
appRegistrationGuid="$3"
appRegistrationApiUri="$4"

# Perform the replacements and write to the new file
sed "s|{{appUri}}|${appUri}|g; s|{{appRegistrationGuid}}|${appRegistrationGuid}|g; s|{{appRegistrationApiUri}}|${appRegistrationApiUri}|g" "$original_xml_file" > "$modified_xml_file"

echo "Replacements completed. Modified XML file saved as $modified_xml_file."


