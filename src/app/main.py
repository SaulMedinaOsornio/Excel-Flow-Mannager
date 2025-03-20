"""
Este script importa la función 'ex' desde el módulo 'gui' y la ejecuta
si el archivo es ejecutado directamente.

Flujo del programa:
1. Se importa la función 'ex' desde el módulo 'gui'.
2. Se verifica si el archivo es ejecutado directamente usando la condición '__name__ == "__main__"'.
3. Si el script es ejecutado directamente, se llama a la función 'ex()'.
"""

# Importación del módulo 'ex' desde el paquete 'gui'
from gui import ex

# Verifica si el script es ejecutado directamente
if __name__ == "__main__":
    """
    Si el script es ejecutado directamente (y no importado),
    se llama a la función 'ex()' definida en el módulo 'gui'.
    """
    ex()
