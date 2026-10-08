txt = open('app.py').read()
txt = txt.replace(
    'status_plantilla = "✅ Cargada" if plantilla_path.exists() else "❌ No cargada"',
    'status_plantilla = f"✅ {plantilla_path.name}" if plantilla_path.exists() else "❌ No cargada"'
)
txt = txt.replace(
    'status_estilos = "✅ Cargado" if estilos_path.exists() else "❌ No cargado"',
    'status_estilos = f"✅ {estilos_path.name}" if estilos_path.exists() else "❌ No cargado"'
)
open('app.py','w').write(txt)
print('OK')
