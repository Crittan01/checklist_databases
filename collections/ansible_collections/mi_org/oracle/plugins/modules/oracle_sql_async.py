#!/usr/bin/python
# -*- coding: utf-8 -*-

DOCUMENTATION = '''
---
module: oracle_sql_async
short_description: Execute Oracle SQL with threading support
description:
    - Execute arbitrary sql against an Oracle database
    - Support for parallel execution of multiple SQL statements using threading
version_added: "2.1.0.0"
options:
    username:
        description:
            - The database username to connect to the database
        required: false
        default: None
        aliases: ['un', 'user']
    password:
        description:
            - The password to connect to the database
        required: false
        default: None
        aliases: ['pw']
    service_name:
        description:
            - The service_name to connect to the database
        required: false
        aliases: ['sn']
    hostname:
        description:
            - The host of the database
        required: false
        default: localhost
        aliases: ['host']
    port:
        description:
            - The listener port to connect to the database
        required: false
        default: 1521
    mode:
        description:
            - Connection mode
        required: false
        default: normal
        choices: ["sysasm", "sysdba", "normal"]
    sql:
        description:
            - Single SQL statement to execute
        required: false
    sql_list:
        description:
            - List of SQL statements to execute in parallel
        required: false
        type: list
    max_workers:
        description:
            - Maximum number of parallel threads
        required: false
        default: 10
        type: int
    timeout:
        description:
            - Timeout in seconds for each SQL statement
        required: false
        default: 300
        type: int
notes:
    - cx_Oracle needs to be installed
    - Oracle client libraries need to be installed
requirements: [ "cx_Oracle" ]
author: Modified for threading support
'''

EXAMPLES = '''
# Execute single SQL
- oracle_sql_async:
    service_name: pdnast
    mode: sysdba
    sql: 'select count(*) from dba_objects where status = ''INVALID''
- oracle_sql_async:
    service_name: pdnast
    mode: sysdba
    sql_list:
      - 'alter view schema.view1 compile'
      - 'alter package schema.pkg1 compile'
      - 'alter function schema.func1 compile'
    max_workers: 20
    timeout: 300
'''

import os
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from ansible.module_utils.basic import AnsibleModule

try:
    import cx_Oracle
    cx_oracle_exists = True
except ImportError:
    cx_oracle_exists = False


class OracleConnectionManager:
    """Manages Oracle connections for threading"""
    
    def __init__(self, connection_params):
        self.connection_params = connection_params
        
    def create_connection(self):
        """Create a new Oracle connection"""
        user = self.connection_params.get('user')
        password = self.connection_params.get('password')
        service_name = self.connection_params.get('service_name')
        hostname = self.connection_params.get('hostname', 'localhost')
        port = self.connection_params.get('port', 1521)
        mode = self.connection_params.get('mode', 'normal')
        
        wallet_connect = '/@%s' % service_name
        
        try:
            if not user and not password:  # Oracle wallet
                if mode == 'sysdba':
                    conn = cx_Oracle.connect(wallet_connect, mode=cx_Oracle.SYSDBA)
                elif mode == 'sysasm':
                    conn = cx_Oracle.connect(wallet_connect, mode=cx_Oracle.SYSASM)
                else:
                    conn = cx_Oracle.connect(wallet_connect)
            elif user and password:
                dsn = cx_Oracle.makedsn(host=hostname, port=port, service_name=service_name)
                if mode == 'sysdba':
                    conn = cx_Oracle.connect(user, password, dsn, mode=cx_Oracle.SYSDBA)
                elif mode == 'sysasm':
                    conn = cx_Oracle.connect(user, password, dsn, mode=cx_Oracle.SYSASM)
                else:
                    conn = cx_Oracle.connect(user, password, dsn)
            else:
                raise Exception('Missing username or password for cx_Oracle')
                
            return conn
            
        except cx_Oracle.DatabaseError as exc:
            error, = exc.args
            raise Exception(f'Could not connect to database: {error.message}')


def execute_single_sql_thread(connection_manager, sql_statement, timeout, thread_index):
    """Execute a single SQL statement in a thread"""
    result = {
        'sql': sql_statement[:100] + '...' if len(sql_statement) > 100 else sql_statement,
        'full_sql': sql_statement,
        'success': False,
        'message': '',
        'error': None,
        'execution_time': 0,
        'thread_id': f'worker-{thread_index}',
        'thread_name': threading.current_thread().name
    }
    
    start_time = time.time()
    conn = None
    cursor = None
    
    try:
        # Create connection for this thread
        conn = connection_manager.create_connection()
        cursor = conn.cursor()
        
        # Set session info for monitoring
        try:
            cursor.execute(
                "BEGIN DBMS_APPLICATION_INFO.SET_MODULE('ansible_async', :1); END;", 
                [f'worker-{thread_index}']
            )
        except:
            pass  # Ignore if DBMS_APPLICATION_INFO is not available
        
        # Execute the SQL
        sql_clean = sql_statement.strip().rstrip(';')
        
        if sql_clean.lower().startswith('select'):
            cursor.execute(sql_clean)
            fetch_result = cursor.fetchall()
            result['message'] = fetch_result
            result['success'] = True
            result['row_count'] = len(fetch_result) if fetch_result else 0
        else:
            cursor.execute(sql_clean)
            conn.commit()
            result['message'] = f'Statement executed successfully'
            result['success'] = True
            result['row_count'] = cursor.rowcount if hasattr(cursor, 'rowcount') else 0
            
    except cx_Oracle.DatabaseError as exc:
        error, = exc.args
        result['error'] = f'Oracle Error {error.code}: {error.message}'
        result['success'] = False
        
    except Exception as e:
        result['error'] = f'General Error: {str(e)}'
        result['success'] = False
        
    finally:
        execution_time = time.time() - start_time
        result['execution_time'] = round(execution_time, 3)
        
        try:
            if cursor:
                cursor.close()
            if conn:
                conn.close()
        except:
            pass  # Ignore cleanup errors
            
    return result


def execute_sql_list_parallel(module, connection_params, sql_list, max_workers, timeout):
    """Execute multiple SQL statements in parallel using ThreadPoolExecutor"""
    
    if not sql_list or len(sql_list) == 0:
        return {
            'summary': {
                'total_statements': 0,
                'successful': 0,
                'failed': 0,
                'success_rate': 0,
                'total_execution_time': 0,
                'max_workers_used': 0,
                'average_time_per_statement': 0
            },
            'detailed_results': [],
            'failed_statements': [],
            'successful_statements': []
        }
    
    connection_manager = OracleConnectionManager(connection_params)
    results = []
    successful = 0
    failed = 0
    
    start_time = time.time()
    
    # Adjust max_workers based on list size
    # actual_workers = min(max_workers, len(sql_list), 50)  # Cap at 50

    if len(sql_list) < 100:
        actual_workers = min(10, max_workers)
    elif len(sql_list) < 500:
        actual_workers = min(15, max_workers)
    elif len(sql_list) < 1000:
        actual_workers = min(20, max_workers)
    elif len(sql_list) < 3000:
        actual_workers = min(30, max_workers)
    else:
        actual_workers = min(50, max_workers)  # Cap óptimo para listas grandes

    try:
        # Use ThreadPoolExecutor for parallel execution
        with ThreadPoolExecutor(max_workers=actual_workers) as executor:
            # Submit all tasks with thread index
            future_to_sql = {
                executor.submit(execute_single_sql_thread, connection_manager, sql, timeout, i): (sql, i)
                for i, sql in enumerate(sql_list)
            }
            
            # Collect results as they complete
            completed = 0
            for future in as_completed(future_to_sql, timeout=timeout*2):  # Overall timeout
                try:
                    result = future.result(timeout=timeout)
                    results.append(result)
                    completed += 1
                    
                    if result['success']:
                        successful += 1
                    else:
                        failed += 1
                        
                    # Progress indicator for large batches
                    if len(sql_list) > 50 and completed % 50 == 0:
                        module.log(f"Completed {completed}/{len(sql_list)} statements")
                        
                except Exception as exc:
                    sql, thread_idx = future_to_sql[future]
                    failed_result = {
                        'sql': sql[:100] + '...' if len(sql) > 100 else sql,
                        'full_sql': sql,
                        'success': False,
                        'message': f'Thread execution failed: {str(exc)}',
                        'error': str(exc),
                        'execution_time': 0,
                        'thread_id': f'worker-{thread_idx}',
                        'thread_name': 'failed'
                    }
                    results.append(failed_result)
                    failed += 1
                    completed += 1
                    
    except Exception as e:
        module.fail_json(msg=f"ThreadPoolExecutor failed: {str(e)}")
    
    total_time = time.time() - start_time
    
    # Sort results by original order if possible
    try:
        results.sort(key=lambda x: int(x['thread_id'].split('-')[1]) if 'worker-' in x['thread_id'] else 999)
    except:
        pass  # Keep original order if sorting fails
    
    # Prepare summary
    summary = {
        'total_statements': len(sql_list),
        'successful': successful,
        'failed': failed,
        'success_rate': round((successful / len(sql_list)) * 100, 2) if len(sql_list) > 0 else 0,
        'total_execution_time': round(total_time, 2),
        'max_workers_used': actual_workers,
        'average_time_per_statement': round(total_time / len(sql_list), 3) if len(sql_list) > 0 else 0,
        'theoretical_sequential_time': round(sum(r['execution_time'] for r in results), 2),
        'speedup_factor': round(sum(r['execution_time'] for r in results) / total_time, 2) if total_time > 0 else 0
    }
    
    return {
        'summary': summary,
        'detailed_results': results,
        'failed_statements': [r for r in results if not r['success']],
        'successful_statements': [r for r in results if r['success']]
    }


def execute_single_sql_sync(module, connection_params, sql):
    """Execute single SQL statement synchronously"""
    connection_manager = OracleConnectionManager(connection_params)
    
    try:
        conn = connection_manager.create_connection()
        cursor = conn.cursor()
        
        sql_clean = sql.strip().rstrip(';')
        
        if sql_clean.lower().startswith('select'):
            cursor.execute(sql_clean)
            result = cursor.fetchall()
            cursor.close()
            conn.close()
            return result
        else:
            cursor.execute(sql_clean)
            conn.commit()
            cursor.close()
            conn.close()
            return f'SQL executed successfully: {sql_clean[:50]}...'
            
    except cx_Oracle.DatabaseError as exc:
        error, = exc.args
        module.fail_json(msg=f'Oracle Error {error.code}: {error.message}', changed=False)
    except Exception as e:
        module.fail_json(msg=f'Error: {str(e)}', changed=False)


def main():
    module = AnsibleModule(
        argument_spec=dict(
            user=dict(required=False, aliases=['un', 'username']),
            password=dict(required=False, no_log=True, aliases=['pw']),
            mode=dict(default="normal", choices=["sysasm", "sysdba", "normal"]),
            service_name=dict(required=False, aliases=['sn']),
            hostname=dict(required=False, default='localhost', aliases=['host']),
            port=dict(required=False, default=1521, type='int'),
            sql=dict(required=False),
            sql_list=dict(required=False, type='list'),
            max_workers=dict(required=False, default=10, type='int'),
            timeout=dict(required=False, default=300, type='int'),
        ),
        mutually_exclusive=[['sql', 'sql_list']],
        required_one_of=[['sql', 'sql_list']]
    )

    if not cx_oracle_exists:
        module.fail_json(msg="The cx_Oracle module is required. Also set LD_LIBRARY_PATH & ORACLE_HOME")

    # Extract parameters
    connection_params = {
        'user': module.params["user"],
        'password': module.params["password"],
        'mode': module.params["mode"],
        'service_name': module.params["service_name"],
        'hostname': module.params["hostname"],
        'port': module.params["port"]
    }
    
    sql = module.params["sql"]
    sql_list = module.params["sql_list"]
    max_workers = module.params["max_workers"]
    timeout = module.params["timeout"]

    # Validate parameters
    if max_workers < 1:
        max_workers = 1
    elif max_workers > 50:
        max_workers = 50

    # Execute SQL
    if sql:
        # Single SQL execution
        result = execute_single_sql_sync(module, connection_params, sql)
        if sql.lower().strip().startswith('select'):
            module.exit_json(msg=result, changed=False)
        else:
            module.exit_json(msg=result, changed=True)
            
    elif sql_list:
        # Parallel SQL execution
        if len(sql_list) == 0:
            module.fail_json(msg="sql_list cannot be empty")
            
        # Execute in parallel
        result = execute_sql_list_parallel(module, connection_params, sql_list, max_workers, timeout)
        
        # Determine if any changes were made (non-SELECT statements)
        changed = any(
            not r['full_sql'].lower().strip().startswith('select') 
            for r in result['successful_statements']
        )
        
        module.exit_json(
            msg=result,
            changed=changed,
            summary=result['summary'],
            failed_count=result['summary']['failed'],
            successful_count=result['summary']['successful']
        )
    
    module.fail_json(msg="No SQL provided")


if __name__ == '__main__':
    main()