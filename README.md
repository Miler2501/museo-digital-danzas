# Museo Digital Danzas - Generador

Este repositorio contiene un script para generar `Museo_Digital_Danzas_Peruanas.pptx` de forma local.

Cómo usar:
1. Instala Python 3 y pip.
2. Instala python-pptx:
   ```
   python3 -m pip install python-pptx
   ```
3. Descarga `create_presentation.py` y ejecútalo:
   ```
   python3 create_presentation.py
   ```
4. Se generará `Museo_Digital_Danzas_Peruanas.pptx` en la misma carpeta. Ábrelo con PowerPoint y ajusta los videos o textos si lo deseas.

Notas:
- El script añade hipervínculos internos y botones "Ver video" que abren enlaces de YouTube en el navegador. No crea "Presentaciones personalizadas" por limitaciones de la librería.
- Si quieres que suba el .pptx al repo por ti, házmelo saber (pero ten en cuenta que subir binarios puede tardar y no siempre es lo ideal).

--
Automation: Copilot
