from pyrevit import revit, DB, forms
import json
import subprocess
import sys

def create_structural_elements_json():
    doc = revit.doc
    output = {}

    categories = {
        "Foundations": DB.BuiltInCategory.OST_StructuralFoundation,
        "Columns": DB.BuiltInCategory.OST_StructuralColumns,
        "Beams": DB.BuiltInCategory.OST_StructuralFraming,
        "Floors": DB.BuiltInCategory.OST_Floors
    }

    for cat_name, cat_enum in categories.items():
        cat_data = []
        elements = DB.FilteredElementCollector(doc).OfCategory(cat_enum).WhereElementIsNotElementType().ToElements()

        for el in elements[:20]:
            el_dict = {"name": "Unknown", "volume_m3": "Unknown", "material": "Unknown"}

        
            try:
                el_dict["name"] = el.Symbol.Family.Name
            except:
                try:
                    el_dict["name"] = "Element_" + str(el.Id.IntegerValue)
                except:
                    el_dict["name"] = "Unknown"

            
            try:
                vol_param = el.get_Parameter(DB.BuiltInParameter.HOST_VOLUME_COMPUTED)
                if vol_param:
                    el_dict["volume_m3"] = vol_param.AsDouble() * 0.0283168
            except:
                pass

            
            try:
                mat_ids = el.GetMaterialIds(False)
                if mat_ids:
                    mat_list = []
                    for mid in mat_ids:
                        mat_elem = doc.GetElement(mid)
                        if mat_elem:
                            mat_list.append(mat_elem.Name)
                    if mat_list:
                        el_dict["material"] = ", ".join(mat_list)
            except:
                pass

            cat_data.append(el_dict)

        output[cat_name] = cat_data


    preview = {}
    for k, v in output.items():
        preview[k] = v[:5]

    with open(r"C:\Users\PS\AppData\Roaming\pyRevit\Extensions\mytools.extension\mytools.tab\quantities.panel\auto-boq.pushbutton\temp.json", "w") as f:
        json.dump(preview, f)

python_path = "python"


script_path = r"c:\Users\PS\AppData\Roaming\pyRevit\Extensions\mytools.extension\mytools.tab\quantities.panel\auto-boq.pushbutton\script2.py"

create_structural_elements_json()
p = subprocess.Popen([python_path, script_path],stdout=subprocess.PIPE,
    stderr=subprocess.PIPE
)

out, err = p.communicate()

print("STDOUT:")
print(out)

print("STDERR:")
print(err)
print("Output:", p.stdout)
print("Errors:", p.stderr)
