# Instrucciones de despliegue remoto

Despues de hacer `commit` y `push` del repositorio local, ejecutar siempre estos pasos en la maquina remota.

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

## Nota para Codex

Este flujo debe aplicarse cada vez que se trabaje en este repositorio y se haya publicado un cambio con `git push`.
