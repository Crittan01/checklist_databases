#!/usr/bin/env python3

DOCUMENTATION = r'''
    name: oracle_excel_inventory
    plugin_type: inventory
    short_description: Oracle Excel inventory source for AWX/AAP
    description: 
        - Reads Oracle database inventory from Excel file
        - Creates localhost with oracle_databases variable structure
        - Designed for opitzconsulting.ansible_oracle.oracle_sql module
    requirements:
        - openpyxl>=3.0.0
    options:
        plugin:
            description: Name of the plugin
            required: true
            choices: ['oracle_excel_inventory']
        excel_file:
            description: Path to Excel file relative to inventory file
            required: true
            type: str
        sheet_name:
            description: Sheet name to read from Excel
            required: false
            default: 'ORACLE'
            type: str
        default_port:
            description: Default Oracle port
            required: false
            default: 1521
            type: int
        validate_data:
            description: Validate Excel data during import
            required: false
            default: true
            type: bool
'''

EXAMPLES = r'''
# oracle_excel_inventory.yml
plugin: oracle_excel_inventory
excel_file: InventarioBD.xlsx
sheet_name: ORACLE
default_port: 1521
validate_data: true
'''

from ansible.plugins.inventory import BaseInventoryPlugin, Constructable, Cacheable
from ansible.errors import AnsibleError, AnsibleParserError
import os

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

class InventoryModule(BaseInventoryPlugin, Constructable, Cacheable):

    NAME = 'oracle_excel_inventory'

    def verify_file(self, path):
        """Verifica que el archivo sea manejable por este plugin"""
        valid = False
        if super(InventoryModule, self).verify_file(path):
            if path.endswith(('oracle_excel_inventory.yml', 'oracle_excel_inventory.yaml')):
                valid = True
        return valid

    def parse(self, inventory, loader, path, cache=True):
        """Parsing principal del inventario"""
        
        if not HAS_OPENPYXL:
            raise AnsibleParserError("openpyxl library is required. Install with: pip install openpyxl>=3.0.0")

        super(InventoryModule, self).parse(inventory, loader, path, cache)
        
        # Leer configuración del archivo YAML
        self._read_config_data(path)
        
        # Obtener configuración
        excel_file = self.get_option('excel_file')
        sheet_name = self.get_option('sheet_name')
        default_port = self.get_option('default_port')
        validate_data = self.get_option('validate_data')
        
        # Resolver path del Excel
        inventory_dir = os.path.dirname(path)
        excel_path = os.path.join(inventory_dir, excel_file)
        
        # Validaciones iniciales
        self._validate_excel_file(excel_path)
        
        # Procesar Excel y generar inventario
        oracle_databases = self._process_excel_file(excel_path, sheet_name, default_port, validate_data)
        
        # Crear inventario simple: solo localhost
        self._create_inventory(oracle_databases)
        
        self.display.v(f"✅ Loaded {len(oracle_databases)} Oracle databases from {excel_file}")

    def _validate_excel_file(self, excel_path):
        """Validaciones de seguridad y existencia del archivo"""
        if not os.path.exists(excel_path):
            raise AnsibleParserError(f"ERROR: Excel file not found: {excel_path}")
        
        if not os.access(excel_path, os.R_OK):
            raise AnsibleParserError(f"ERROR: Cannot read Excel file: {excel_path}")
        
        # Validar tamaño (máximo 50MB para evitar problemas de memoria)
        file_size = os.path.getsize(excel_path) / (1024 * 1024)  # MB
        if file_size > 50:
            self.display.warning(f"Large Excel file ({file_size:.1f}MB). Consider optimizing.")

    def _create_inventory(self, oracle_databases):
        """Crea la estructura de inventario simple"""
        # Solo localhost como host
        self.inventory.add_host('localhost')
        
        # Variables de conexión para localhost
        self.inventory.set_variable('localhost', 'ansible_connection', 'local')
        self.inventory.set_variable('localhost', 'ansible_python_interpreter', '{{ ansible_playbook_python }}')
        
        # La estructura oracle_databases como variable de localhost
        self.inventory.set_variable('localhost', 'oracle_databases', oracle_databases)
        
        # Variables globales útiles para reportes
        self.inventory.set_variable('all', 'oracle_database_count', len(oracle_databases))
        self.inventory.set_variable('all', 'inventory_source', 'oracle_excel')
        self.inventory.set_variable('all', 'inventory_timestamp', self._get_timestamp())
        
        # Estadísticas por ambiente
        env_stats = self._calculate_environment_stats(oracle_databases)
        self.inventory.set_variable('all', 'oracle_environments', env_stats)

    def _process_excel_file(self, excel_path, sheet_name, default_port, validate_data):
        """Procesa el archivo Excel y retorna la estructura oracle_databases"""
        
        oracle_databases = {}
        
        try:
            # Cargar Excel optimizado para lectura
            workbook = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
            
            # Verificar que la hoja existe
            if sheet_name not in workbook.sheetnames:
                available_sheets = ", ".join(workbook.sheetnames)
                raise AnsibleParserError(f"ERROR: Sheet '{sheet_name}' not found. Available sheets: {available_sheets}")
            
            worksheet = workbook[sheet_name]
            
            # Procesar headers
            headers = self._extract_headers(worksheet)
            column_mapping = self._get_column_mapping(headers)
            
            # Validar columnas requeridas
            if validate_data:
                self._validate_required_columns(column_mapping)
            
            # Procesar filas de datos
            processed_count = 0
            error_count = 0
            
            for row_num, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=2):
                if not any(row):  # Saltar filas completamente vacías
                    continue
                
                try:
                    db_config = self._process_excel_row(row, column_mapping, default_port, validate_data)
                    if db_config:
                        instancia = db_config['instancia']
                        
                        # Verificar duplicados
                        if instancia in oracle_databases:
                            self.display.warning(f"Duplicate instance '{instancia}' in row {row_num}, skipping")
                            continue
                        
                        oracle_databases[instancia] = db_config
                        processed_count += 1
                        
                except Exception as e:
                    error_count += 1
                    self.display.warning(f"Error in row {row_num}: {str(e)}")
                    continue
            
            workbook.close()
            
            # Reporte final
            self.display.v(f"📊 Processed {processed_count} databases, {error_count} errors")
            
            if not oracle_databases:
                raise AnsibleParserError("ERROR: No valid Oracle databases found in Excel file")
            
        except Exception as e:
            if isinstance(e, AnsibleParserError):
                raise
            raise AnsibleParserError(f"ERROR: Reading Excel file {excel_path}: {str(e)}")
        
        return oracle_databases

    def _extract_headers(self, worksheet):
        """Extrae y limpia headers de la primera fila"""
        headers = []
        for cell in worksheet[1]:
            header = str(cell.value).strip() if cell.value else ''
            headers.append(header)
        return headers

    def _get_column_mapping(self, headers):
        """Mapea las columnas del Excel a campos esperados"""
        mapping = {}
        
        # Mapeo flexible de nombres de columnas
        column_names = {
            'instancia': ['Instancia', 'Instance', 'instancia', 'INSTANCIA'],
            'host': ['Host', 'hostname', 'server', 'HOST'],
            'direccion_ip': ['Dirección IP', 'IP Address', 'IP', 'direccion_ip', 'Direccion_IP'],
            'cadena_tns': ['Cadena TNS', 'TNS', 'tns_name', 'cadena_tns', 'CADENA_TNS'],
            'motor_bd': ['Motor BD', 'Database Engine', 'engine', 'MOTOR_BD'],
            'sistema_operativo': ['Sistema Operativo', 'OS', 'operating_system', 'SISTEMA_OPERATIVO'],
            'release_so': ['Release  SO', 'Release SO', 'OS Release', 'os_version', 'RELEASE_SO'],
            'release_bd': ['Release BD', 'DB Release', 'db_version', 'RELEASE_BD'],
            'ambiente': ['Ambiente', 'Environment', 'env', 'AMBIENTE'],
            'datacenter': ['Datacenter', 'DC', 'datacenter', 'DATACENTER'],
            'compania': ['Compañía', 'Company', 'compania', 'COMPANIA'],
            'transversal': ['Transversal', 'transversal', 'TRANSVERSAL'],
            'fecha_creacion': ['Fecha Creación', 'Creation Date', 'created', 'FECHA_CREACION'],
            'cpu': ['CPU', 'cpu', 'CPUs'],
            'ram': ['RAM', 'Memory', 'ram', 'MEMORIA'],
            'disco_gb': ['Disco GB', 'Disk GB', 'storage', 'DISCO_GB'],
            'ip_nat': ['IP NAT', 'NAT IP', 'ip_nat', 'IP_NAT']
        }
        
        # Buscar coincidencias (case-sensitive first, then insensitive)
        for field, possible_names in column_names.items():
            for i, header in enumerate(headers):
                if header in possible_names:
                    mapping[field] = i
                    break
        
        return mapping

    def _validate_required_columns(self, column_mapping):
        """Valida que las columnas requeridas estén presentes"""
        required_columns = ['instancia']
        missing_columns = []
        
        for col in required_columns:
            if col not in column_mapping:
                missing_columns.append(col)
        
        if missing_columns:
            raise AnsibleParserError(f"ERROR: Missing required columns: {', '.join(missing_columns)}")

    def _process_excel_row(self, row, column_mapping, default_port, validate_data):
        """Procesa una fila del Excel y retorna configuración de BD"""
        
        # Obtener instancia (campo obligatorio)
        if 'instancia' not in column_mapping:
            raise AnsibleError("Column 'Instancia' not found")
        
        instancia = row[column_mapping['instancia']]
        if not instancia or str(instancia).strip() == '':
            return None  # Saltar filas sin instancia
        
        instancia = str(instancia).strip()
        
        # Función helper para obtener valores
        def get_col_value(field, default=''):
            if field in column_mapping and column_mapping[field] < len(row):
                value = row[column_mapping[field]]
                if value is not None and str(value).strip() != '':
                    return str(value).strip()
            return default
        
        # Validaciones si está habilitada
        if validate_data:
            if not get_col_value('direccion_ip'):
                self.display.warning(f"Missing IP address for instance '{instancia}'")
        
        # Construir configuración de base de datos (SIN credenciales)
        db_config = {
            # 🎯 Datos de conexión (sin usuario/password)
            'instancia': instancia,
            'host': get_col_value('host', instancia.lower()),
            'direccion_ip': get_col_value('direccion_ip'),
            'cadena_tns': get_col_value('cadena_tns', instancia),
            'puerto': default_port,
            
            # 📊 Metadatos para reporte y análisis
            'motor_bd': get_col_value('motor_bd', 'Oracle'),
            'sistema_operativo': get_col_value('sistema_operativo'),
            'release_so': get_col_value('release_so'),
            'release_bd': get_col_value('release_bd'),
            'ambiente': get_col_value('ambiente'),
            'datacenter': get_col_value('datacenter'),
            'compania': get_col_value('compania'),
            'transversal': get_col_value('transversal'),
            'fecha_creacion': get_col_value('fecha_creacion'),
            
            # 💾 Recursos de servidor
            'cpu': self._safe_int(get_col_value('cpu', '0')),
            'ram': self._safe_int(get_col_value('ram', '0')),
            'disco_gb': self._safe_int(get_col_value('disco_gb', '0')),
            
            # 🌐 Networking opcional
            'ip_nat': get_col_value('ip_nat') if get_col_value('ip_nat') not in ['-', 'N/A', '#N/A'] else None
        }
        
        return db_config

    def _safe_int(self, value):
        """Convierte valor a int de forma segura"""
        try:
            if str(value).strip() in ['', '-', 'N/A', '#N/A']:
                return 0
            # Manejar valores como "1,024" o "1.5"
            clean_value = str(value).replace(',', '').replace(' ', '')
            return int(float(clean_value))
        except (ValueError, TypeError):
            return 0

    def _get_timestamp(self):
        """Obtiene timestamp actual"""
        from datetime import datetime
        return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    def _calculate_environment_stats(self, oracle_databases):
        """Calcula estadísticas por ambiente"""
        env_stats = {}
        
        for db_name, db_config in oracle_databases.items():
            ambiente = db_config.get('ambiente', 'Unknown')
            if ambiente not in env_stats:
                env_stats[ambiente] = {
                    'count': 0,
                    'instances': [],
                    'datacenters': set(),
                    'companies': set()
                }
            
            env_stats[ambiente]['count'] += 1
            env_stats[ambiente]['instances'].append(db_name)
            
            if db_config.get('datacenter'):
                env_stats[ambiente]['datacenters'].add(db_config['datacenter'])
            
            if db_config.get('compania'):
                env_stats[ambiente]['companies'].add(db_config['compania'])
        
        # Convertir sets a lists para serialización JSON
        for env in env_stats:
            env_stats[env]['datacenters'] = list(env_stats[env]['datacenters'])
            env_stats[env]['companies'] = list(env_stats[env]['companies'])
        
        return env_stats
#!/usr/bin/env python3

DOCUMENTATION = r'''
    name: oracle_excel_inventory
    plugin_type: inventory
    short_description: Oracle Excel inventory source
    description: Reads Oracle database inventory from Excel file and creates localhost with oracle_databases variable
    requirements:
        - openpyxl>=3.0.0
    options:
        plugin:
            description: Name of the plugin
            required: true
            choices: ['oracle_excel_inventory']
        excel_file:
            description: Path to Excel file relative to inventory file
            required: true
            type: str
        sheet_name:
            description: Sheet name to read from Excel
            required: false
            default: 'Oracle'
            type: str
        default_port:
            description: Default Oracle port
            required: false
            default: 1521
            type: int
        # No necesitamos default_usuario ni default_password
        # Las credenciales se manejan via AWX Credentials
'''

EXAMPLES = r'''
# oracle_excel_inventory.yml
plugin: oracle_excel_inventory
excel_file: InventarioBD.xlsx
sheet_name: Oracle
default_port: 1521
'''

from ansible.plugins.inventory import BaseInventoryPlugin, Constructable, Cacheable
from ansible.errors import AnsibleError, AnsibleParserError
import os

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

class InventoryModule(BaseInventoryPlugin, Constructable, Cacheable):

    NAME = 'oracle_excel_inventory'

    def verify_file(self, path):
        """Verifica que el archivo sea manejable por este plugin"""
        valid = False
        if super(InventoryModule, self).verify_file(path):
            # El archivo debe terminar en oracle_excel_inventory.yml o .yaml
            if path.endswith(('oracle_excel_inventory.yml', 'oracle_excel_inventory.yaml')):
                valid = True
        return valid

    def parse(self, inventory, loader, path, cache=True):
        """Parsing principal del inventario"""
        
        if not HAS_OPENPYXL:
            raise AnsibleParserError("openpyxl library is required for oracle_excel_inventory plugin")

        super(InventoryModule, self).parse(inventory, loader, path, cache)
        
        # Leer configuración del archivo YAML
        self._read_config_data(path)
        
        # Obtener configuración
        excel_file = self.get_option('excel_file')
        sheet_name = self.get_option('sheet_name')
        default_port = self.get_option('default_port')
        
        # Resolver path del Excel relativo al directorio del inventory
        inventory_dir = os.path.dirname(path)
        excel_path = os.path.join(inventory_dir, excel_file)
        
        if not os.path.exists(excel_path):
            raise AnsibleParserError(f"Excel file not found: {excel_path}")
        
        # Procesar Excel y generar inventario
        oracle_databases = self._process_excel_file(excel_path, sheet_name, default_port)
        
        # Crear localhost como único host
        self.inventory.add_host('localhost')
        self.inventory.set_variable('localhost', 'ansible_connection', 'local')
        self.inventory.set_variable('localhost', 'ansible_python_interpreter', 
                                  '{{ ansible_playbook_python }}')
        
        # Asignar la estructura oracle_databases a localhost
        self.inventory.set_variable('localhost', 'oracle_databases', oracle_databases)
        
        # Variables globales útiles
        self.inventory.set_variable('all', 'oracle_database_count', len(oracle_databases))
        self.inventory.set_variable('all', 'inventory_source', 'oracle_excel')
        
        self.display.v(f"Loaded {len(oracle_databases)} Oracle databases from {excel_file}")

    def _process_excel_file(self, excel_path, sheet_name, default_port):
        """Procesa el archivo Excel y retorna la estructura oracle_databases"""
        
        oracle_databases = {}
        
        try:
            # Abrir Excel en modo read-only para mejor performance
            workbook = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
            
            if sheet_name not in workbook.sheetnames:
                raise AnsibleParserError(f"Sheet '{sheet_name}' not found in Excel file. Available sheets: {workbook.sheetnames}")
            
            worksheet = workbook[sheet_name]
            
            # Obtener headers de la primera fila
            headers = []
            for cell in worksheet[1]:
                headers.append(cell.value.strip() if cell.value else '')
            
            # Mapear columnas esperadas
            column_mapping = self._get_column_mapping(headers)
            
            # Procesar cada fila de datos
            for row_num, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=2):
                if not any(row):  # Saltar filas vacías
                    continue
                
                try:
                    db_config = self._process_excel_row(row, column_mapping, default_port)
                    if db_config:
                        # Crear clave única combinando instancia + ambiente + datacenter
                        unique_key = self._create_unique_key(db_config)
                        
                        # Verificar si ya existe y mostrar warning
                        if unique_key in oracle_databases:
                            self.display.warning(f"Duplicate key '{unique_key}' in row {row_num}, overwriting previous entry")
                        
                        oracle_databases[unique_key] = db_config
                        
                except Exception as e:
                    self.display.warning(f"Error processing row {row_num}: {str(e)}")
                    continue
            
            workbook.close()
            
        except Exception as e:
            raise AnsibleParserError(f"Error reading Excel file {excel_path}: {str(e)}")
        
        return oracle_databases

    def _create_unique_key(self, db_config):
        """Crea una clave única para evitar duplicados"""
        instancia = db_config['instancia']
        ambiente = db_config['ambiente'].replace(' ', '_') if db_config['ambiente'] else 'Unknown'
        datacenter = db_config['datacenter'] if db_config['datacenter'] else 'Unknown'
        
        # Crear clave base: INSTANCIA_AMBIENTE_DATACENTER
        base_key = f"{instancia}_{ambiente}_{datacenter}".upper()
        
        # Si es la misma instancia en el mismo ambiente/datacenter pero host diferente,
        # agregar parte del host para diferenciar
        if db_config.get('host'):
            # Tomar últimos 2 caracteres del host para diferenciación
            host_suffix = db_config['host'][-2:] if len(db_config['host']) >= 2 else db_config['host']
            unique_key = f"{base_key}_{host_suffix}".upper()
        else:
            unique_key = base_key
        
        return unique_key

    def _get_column_mapping(self, headers):
        """Mapea las columnas del Excel a campos esperados"""
        mapping = {}
        
        # Mapeo de nombres de columnas (case-insensitive y variaciones)
        column_names = {
            'instancia': ['Instancia', 'Instance', 'instancia'],
            'host': ['Host', 'hostname', 'server'],
            'direccion_ip': ['Direccion_IP', 'Dirección IP', 'IP Address', 'IP', 'direccion_ip'],
            'cadena_tns': ['Cadena_TNS', 'Cadena TNS', 'TNS', 'tns_name', 'cadena_tns'],
            'motor_bd': ['Motor_BD', 'Motor BD', 'Database Engine', 'engine'],
            'sistema_operativo': ['Sistema_Operativo', 'Sistema Operativo', 'OS', 'operating_system'],
            'release_so': ['Release_SO', 'Release  SO', 'OS Release', 'os_version'],
            'release_bd': ['Release_BD', 'Release BD', 'DB Release', 'db_version'],
            'ambiente': ['Ambiente', 'Environment', 'env'],
            'datacenter': ['Datacenter', 'DC', 'datacenter'],
            'compania': ['Companía', 'Compañía', 'Company', 'compania'],
            'transversal': ['Transversal', 'transversal'],
            'fecha_creacion': ['Fecha_Creacion', 'Fecha Creación', 'Creation Date', 'created'],
            'cpu': ['CPU', 'cpu'],
            'ram': ['RAM', 'Memory', 'ram'],
            'disco_gb': ['Disco_GB', 'Disco GB', 'Disk GB', 'storage'],
            'ip_nat': ['IP_NAT', 'IP NAT', 'NAT IP', 'ip_nat']
        }
        
        for field, possible_names in column_names.items():
            for i, header in enumerate(headers):
                if header in possible_names:
                    mapping[field] = i
                    break
        
        return mapping

    def _process_excel_row(self, row, column_mapping, default_port):
        """Procesa una fila del Excel y retorna configuración de BD"""
        
        # Validar que tenemos al menos la instancia
        if 'instancia' not in column_mapping:
            raise AnsibleError("Column 'Instancia' not found in Excel")
        
        instancia = row[column_mapping['instancia']]
        if not instancia:
            return None
        
        # Función helper para obtener valor de columna
        def get_col_value(field, default=''):
            if field in column_mapping and column_mapping[field] < len(row):
                value = row[column_mapping[field]]
                return str(value).strip() if value not in [None, ''] else default
            return default
        
        # Construir configuración de base de datos
        db_config = {
            # Datos específicos para oracle_sql module
            'hostname': get_col_value('direccion_ip', get_col_value('host')),  # Priorizar IP sobre hostname
            'port': default_port,
            'service_name': get_col_value('cadena_tns', str(instancia)),
            
            # Datos originales de la instancia
            'instancia': str(instancia).strip(),
            'host': get_col_value('host', str(instancia).lower()),
            'direccion_ip': get_col_value('direccion_ip'),
            'cadena_tns': get_col_value('cadena_tns', str(instancia)),
            
            # Metadatos para reporte y agrupación
            'motor_bd': get_col_value('motor_bd', 'Oracle'),
            'sistema_operativo': get_col_value('sistema_operativo'),
            'release_so': get_col_value('release_so'),
            'release_bd': get_col_value('release_bd'),
            'ambiente': get_col_value('ambiente'),
            'datacenter': get_col_value('datacenter'),
            'compania': get_col_value('compania'),
            'transversal': get_col_value('transversal'),
            'fecha_creacion': get_col_value('fecha_creacion'),
            
            # Recursos de servidor
            'cpu': self._safe_int(get_col_value('cpu', '0')),
            'ram': self._safe_int(get_col_value('ram', '0')),
            'disco_gb': self._safe_int(get_col_value('disco_gb', '0')),
            
            # IP NAT (opcional)
            'ip_nat': get_col_value('ip_nat') if get_col_value('ip_nat') not in ['-', '#N/A'] else None
        }
        
        return db_config

    def _safe_int(self, value):
        """Convierte valor a int de forma segura"""
        try:
            return int(float(str(value))) if value not in ['', None, '-'] else 0
        except (ValueError, TypeError):
            return 0