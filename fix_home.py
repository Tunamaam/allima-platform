
file_path = 'd:/CUENTOS/CLONACION/home.html'
try:
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Fix template literals
    content = content.replace('$ {', '${')
    
    # Fix HTML tags in JS strings
    content = content.replace('< div', '<div')
    content = content.replace('</div >', '</div>')
    
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Successfully fixed home.html")
except Exception as e:
    print(f"Error: {e}")
