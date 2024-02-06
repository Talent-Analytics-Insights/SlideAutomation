# Slide Automation and Custom Table Styles with Python
Python implementation of automating large slide decks with simple contents. For this release, only a title slide and a slide with a table and optional title are included. Our implementation features creating custom Table Styles based on your own colors. 


## Simple Usage

The `pandas` library is imported as `pd`.

The `ppt_builder` library is imported, which contains the main `PPTBuilder` class.

```python
import pandas as pd
from PythonPPTAutomation.ppt_builder import PPTBuilder
```
 

An instance of the `PPTBuilder` class is created and assigned to the `slide_builder` variable.
```python
slide_builder = PPTBuilder()
```

Optionally, you can load your own template. Then you can automatically create slides in your own theme.

```python
slide_builder.load_template('template.pptx')
```


Create your custom TableStyle based on three colors; the color of the table header, the color of the even rows, and the color of the odd rows. These colors can be the same.

```python
custom_theme = ('EE3E41', '7AC043', '419ED7')
```

Name your custom theme and add to the slide builder.

```python
slide_builder.add_table_style_to_template(custom_theme, "Custom Theme")
```

Now create mock data. You can replace this with your own `pd.DataFrame`.
```python
data = {
    'Column 1': ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'],        
    'Column 2': ['K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T'],        
    'Column 3': ['U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD']}

df = pd.DataFrame(data)
```

Now lets add a title slide. You can specify the template id here, with zero corresponding to the first template (you can check this by opening in PowerPoint and clicking 'New Slide', the templates are then listed).

```python
slide_builder.add_title_only_slide(0, "Automatically Generated Title Slide")
```

Now we can add a slide with our table in our custom Table Style.

```python
slide_builder.add_table_only_slide(df, style_name="Custom Theme", title="Table in Custom Theme")
```

Finally, we can save our generated PowerPoint
```python
slide_builder.save("automated.pptx")
```

This code can reduce the time required to build a slide deck displaying a lot of tables repetitively. It saved us a lot of time, which is why we wanted to share this with you. Find the full code below.

```python
import pandas as pd
from PythonPPTAutomation.ppt_builder import PPTBuilder


slide_builder = PPTBuilder()

custom_theme = ('EE3E41', '7AC043', '419ED7')

slide_builder.load_template('template.pptx')

slide_builder.add_table_style_to_template(custom_theme, "Custom Theme")

data = {'Column 1': ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'],  
        'Column 2': ['K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T'],  
        'Column 3': ['U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD']}  
  
df = pd.DataFrame(data)  

slide_builder.add_title_only_slide(0, "Automatically Generated Title Slide")

slide_builder.add_table_only_slide(df, style_name="Custom Theme", title="Table in Custom Theme")

slide_builder.save("automated.pptx")
```

In future releases, we plan to automate many other slide types, such as different types of graphs, multiple tables on one slide, and text slides.
