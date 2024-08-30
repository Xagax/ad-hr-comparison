
def formatear_nombre(nombre_completo):
    # Check if the comma exists in nombre_completo
    if ", " in nombre_completo:
        # Split the string into parts using ", "
        parts = nombre_completo.split(", ")

        # Extract the apellido (last name)
        apellido = parts[0]

        # Extract the rest as nombre (first name and possibly middle name)
        nombre_parts = parts[1].split()  # Split the first name and possibly middle name
        nombre = nombre_parts[0]  # Take the first name
        # # Optionally, if there are more than one part, join them as well
        # if len(nombre_parts) > 1:
        #     nombre += " " + " ".join(nombre_parts[1:])  # Join the remaining parts as well

    else:
        # If the comma doesn't exist, split using any whitespace character
        words = nombre_completo.split()
        # Take the first word as the nombre
        nombre = words[0]

        # Take the last word as the apellido
        apellido = words[-1]

    # Convertir el nombre y el apellido a mayúsculas
    nombre = nombre.title()
    apellido = apellido.title()

    # Devolver el nombre y el apellido en el formato requerido
    nombre_formateado = nombre+" "+apellido
    
    return nombre_formateado
