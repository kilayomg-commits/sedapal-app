txt = open('app.py').read()
txt = txt.replace(
    'st.markdown(f"**Plantilla vacía (.xlsm)** — {status_plantilla}")',
    'st.markdown(f"**Plantilla vacía**\\n\\n{status_plantilla}")'
)
txt = txt.replace(
    'st.markdown(f"**Archivo de estilos/referencia (.xlsm)** — {status_estilos}")',
    'st.markdown(f"**Referencia/Estilos**\\n\\n{status_estilos}")'
)
open('app.py','w').write(txt)
print('OK')
