# -*- coding: utf-8 -*-
"""Genera app.ico: el logo recortado en circulo, sin esquinas negras.

Usado por el workflow de GitHub Actions antes de compilar con PyInstaller.
"""
import os
from PIL import Image, ImageDraw

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
ORIGEN = os.path.join(RAIZ, "Logo.png")
DESTINO = os.path.join(RAIZ, "app.ico")

imagen = Image.open(ORIGEN).convert("RGB")
ancho, alto = imagen.size
corte = min(ancho, alto)
izq, arriba = (ancho - corte) // 2, (alto - corte) // 2
imagen = imagen.crop((izq, arriba, izq + corte, arriba + corte))

LADO = 1024
imagen = imagen.resize((LADO, LADO), Image.LANCZOS)
mascara = Image.new("L", (LADO, LADO), 0)
ImageDraw.Draw(mascara).ellipse((0, 0, LADO - 1, LADO - 1), fill=255)
imagen = imagen.convert("RGBA")
imagen.putalpha(mascara)

imagen.save(DESTINO, format="ICO", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print(f"icono generado: {DESTINO}")
