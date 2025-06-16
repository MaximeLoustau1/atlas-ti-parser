# Atlas-Ti parser that outputs that outputs a formatted xlsx file of groupes codes

# Before parsing:

Make sure each code value is annotated by 'Tn' (n = code number) at the end. 
Here is an example of valid project:


# Parsing:

- Export atlas.ti to xml 
- Run `pip install -r requirements.txt` in the terminal (ideally create a python virtual environment)
- Update `name_of_paper_xml` variable to parse the correct file
- Run main.py and open output.xlsx for results