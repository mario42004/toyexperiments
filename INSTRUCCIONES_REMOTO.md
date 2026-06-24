# Instrucciones de despliegue remoto rapido

Despues de hacer `commit` y `push` del repositorio local, ejecutar siempre estos pasos en la maquina remota. Evitar repetir comprobaciones largas si no hay errores: este repo ya quedo preparado para levantar Streamlit y ngrok.

## Datos remotos

- Host: `10.8.0.113`
- Usuario: `labgpu`
- Directorio remoto: `/home/labgpu/Descargas/unidades/mario_gpu/tesis_2026/toyexperiments`
- Puerto de Streamlit: `8501`

Las credenciales estan en `Script/1 K medoides/credentials.txt`. No copiar passwords en logs, commits ni mensajes.

## Flujo obligatorio

1. Conectarse por SSH a la maquina remota.
2. Entrar en el directorio remoto:

   ```bash
   cd /home/labgpu/Descargas/unidades/mario_gpu/tesis_2026/toyexperiments
   ```

3. Actualizar el repositorio remoto:

   ```bash
   git pull
   ```

4. Lanzar el script de Streamlit en el puerto `8501`.
5. Lanzar ngrok apuntando al puerto de Streamlit:

   ```bash
   ngrok http 8501
   ```

6. Compartir la URL publica de ngrok con el usuario.

## Ruta rapida para Codex

Usar las credenciales de `Script/1 K medoides/credentials.txt` y ejecutar los comandos remotos directamente. `python3-pip`, `streamlit` y `openpyxl` ya fueron instalados en la maquina remota el 2026-06-24; no reinstalar dependencias salvo que Streamlit falle.

Comando remoto recomendado despues del `push`:

```bash
cd /home/labgpu/Descargas/unidades/mario_gpu/tesis_2026/toyexperiments
git pull
pgrep -f "[s]treamlit run Script/1 K medoides/k_medoids_streamlit.py" | xargs -r kill
setsid python3 -m streamlit run "Script/1 K medoides/k_medoids_streamlit.py" --server.port 8501 --server.address 0.0.0.0 --server.headless true > streamlit.log 2>&1 < /dev/null &
```

Despues levantar ngrok si no esta ya corriendo:

```bash
pgrep -af "[n]grok http 8501" || setsid ngrok http 8501 > ngrok.log 2>&1 < /dev/null &
```

Esperar unos segundos y obtener la URL publica con:

```bash
python3 - <<'PY'
import json, urllib.request
data = json.load(urllib.request.urlopen("http://127.0.0.1:4040/api/tunnels", timeout=5))
for tunnel in data.get("tunnels", []):
    print(tunnel.get("public_url", ""))
PY
```

Si el API de ngrok no responde, revisar `ngrok.log`. Para evitar falsos positivos al buscar procesos, usar siempre patrones como `[n]grok http 8501` y `[s]treamlit run ...`, no el texto literal directamente con `pgrep` o `pkill`.

## Nota para Codex

Este flujo debe aplicarse cada vez que se trabaje en este repositorio y se haya publicado un cambio con `git push`.
