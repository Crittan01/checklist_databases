#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
Oracle Inventory Plugin for Ansible - Minimalista
Genera estructura oracle_databases desde Excel para proyecto Oracle Checklist
"""

from __future__ import (absolute_import, division, print_function)
__metaclass__ = type

DOCUMENTATION = '''
---
name: oracle_inventory
author: Oracle Team
version_added: "1.0.0"
short_description: Plugin específico para inventario Oracle desde Excel
description:
    - Genera estructura oracle_databases desde Excel
    - Diseñado específicamente para proyecto Oracle Checklist
    - Mantiene arquitectura instancia-centric
requirements:
    - python >= 3.6
    - pandas >= 1.0.0
    - openpyxl >= 3.0.0

options:
    plugin:
        description: Identificador del plugin
        type: str
        required: true
        choices: ['oracle_inventory']
        
    excel_file:
        description: Ruta al archivo Excel
        type: str
        required: true
        
    sheet_name:
        description: Nombre de la hoja Excel
        type: str
        default: "ORACLE"
        
    filters:
        description: Filtros para aplicar
        type: dict
        default: {}
'''

EXAMPLES = '''
# Configuración básica
---
plugin: oracle_inventory
excel_file: "inventory/InventarioBD.xlsx"
sheet_name: "ORACLE"

# Con filtros
---
plugin: oracle_inventory
excel_file: "inventory/InventarioBD.xlsx"
sheet_name: "ORACLE"
filters:
  Ambiente: "Producción"
  Host: "sgorcbdp03"
'''

import os
import sys
from ansible.plugins.inventory import BaseInventoryPlugin
from ansible.errors import AnsibleError, AnsibleParserError
from ansible.utils.display import Display

display = Display()

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False


class InventoryModule(BaseInventoryPlugin):
    """Plugin minimalista para inventario Oracle desde Excel"""

    NAME = 'oracle_inventory'

    def verify_file(self, path):
        """Verificar que el archivo es válido para este plugin"""
        if super().verify_file(path):
            if path.endswith(('oracle_inventory.yaml', 'oracle_inventory.yml')):
                return True
            raise AnsibleParserError(
                "El archivo debe terminar en .oracle_inventory.yml o .oracle_inventory.yaml"
            )
        return False

    def parse(self, inventory, loader, path, cache=True):
        """Parsear el inventario y generar oracle_databases"""
        super().parse(inventory, loader, path, cache)

        # Verificar dependencias
        if not HAS_PANDAS:
            raise AnsibleError(
                "Se requiere 'pandas' para leer Excel. "
                "Instala con: pip install pandas openpyxl"
            )

        # Leer configuración
        config = self._read_config_data(path)
        
        excel_file = config.get('excel_file')
        sheet_name = config.get('sheet_name', 'ORACLE')
        filters = config.get('filters', {})

        if not excel_file:
            raise AnsibleParserError("Falta parámetro requerido: excel_file")

        # Verificar archivo Excel
        if not os.path.exists(excel_file):
            raise AnsibleError(f"Archivo Excel no encontrado: {excel_file}")

        # Leer y procesar Excel
        oracle_databases = self._process_excel(excel_file, sheet_name, filters)

        # Generar inventario con localhost
        self.inventory.add_host('localhost')
        
        # Establecer oracle_databases como variable de inventario
        self.inventory.set_variable('localhost', 'oracle_databases', oracle_databases)
        
        # Crear grupos por ambiente
        self._create_groups(oracle_databases)

        display.v(f"Inventario Oracle generado: {len(oracle_databases)} instancias")

    def _process_excel(self, excel_file, sheet_name, filters):
        """Procesar archivo Excel y generar oracle_databases"""
        
        try:
            # Leer Excel
            df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
            display.vv(f"Excel leído: {len(df)} filas, {len(df.columns)} columnas")
            
        except Exception as e:
            raise AnsibleError(f"Error leyendo Excel: {str(e)}")

        # Validar columnas requeridas
        required_columns = ['BasedeDatos', 'Ambiente', 'Host', 'DireccionIP']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            available_columns = ', '.join(df.columns.tolist())
            raise AnsibleParserError(
                f"Columnas faltantes: {missing_columns}. "
                f"Disponibles: {available_columns}"
            )

        # Aplicar filtros
        if filters:
            for column, value in filters.items():
                if value and column in df.columns:
                    df = df[df[column].astype(str).str.strip() == str(value).strip()]
                    display.vv(f"Filtro aplicado {column}={value}: {len(df)} filas")

        # Generar oracle_databases
        oracle_databases = {}
        
        for _, row in df.iterrows():
            base_datos = str(row['BasedeDatos']).strip()
            
            if not base_datos or base_datos.lower() in ['nan', 'none']:
                continue

            # Generar estructura compatible con proyecto actual
            oracle_databases[base_datos] = {
                'instancia': base_datos,
                'host': str(row['Host']).strip(),
                'direccion_ip': str(row['DireccionIP']).strip(),
                'cadena_tns': base_datos,
                'puerto': self._extract_port(row),
                'usuario': 'ansible_user',
                'password': 'ansibledb123*',
                'motor_bd': 'Oracle',
                'sistema_operativo': 'Red Hat Enterprise Linux Server',
                'release_so': '8.x',
                'release_bd': '19.0.0.0',
                'ambiente': str(row['Ambiente']).strip(),
                # Campos adicionales del Excel
                'scan': row.get('SCAN', ''),
                'conexion': row.get('CONEXIÓN', '')
            }

        display.v(f"oracle_databases generado con {len(oracle_databases)} instancias")
        return oracle_databases

    def _extract_port(self, row):
        """Extraer puerto de la fila"""
        # Intentar obtener puerto del Excel
        if 'PUERTO' in row:
            puerto_value = str(row['PUERTO']).strip()
            if puerto_value and puerto_value != 'nan':
                # Limpiar formato (quitar comas)
                puerto_clean = puerto_value.replace(',', '').replace('.', '')
                try:
                    return int(puerto_clean)
                except ValueError:
                    pass
        
        # Puerto por defecto
        return 1537

    def _create_groups(self, oracle_databases):
        """Crear grupos por ambiente"""
        grupos_ambiente = {}
        
        for db_key, db_info in oracle_databases.items():
            ambiente = db_info['ambiente'].lower().replace(' ', '_')
            grupo_name = f"oracle_{ambiente}"
            
            if grupo_name not in grupos_ambiente:
                grupos_ambiente[grupo_name] = []
                self.inventory.add_group(grupo_name)
            
            grupos_ambiente[grupo_name].append(db_key)
            
            # Agregar localhost a todos los grupos (para compatibilidad)
            self.inventory.add_child(grupo_name, 'localhost')

        # Establecer información de grupos como variable
        self.inventory.set_variable('localhost', 'grupos_ambiente', grupos_ambiente)
        
        display.vv(f"Grupos creados: {list(grupos_ambiente.keys())}")