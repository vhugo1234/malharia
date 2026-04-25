"""
VERSION_MANAGER.PY — Gerenciamento de Versões de Arquivos Master
=================================================================

Automatiza backup, versionamento e restauração de arquivos Master CDR.
Mantém histórico completo com checksums MD5 para integridade.

Uso:
    vm = VersionManager(output_dir)
    vm.create_backup(master_cdr_path, pedido_id)  # Backup automático
    vm.list_versions(pedido_id)                    # Lista histórico
    vm.restore_version(pedido_id, "v001")          # Restaura versão
    vm.cleanup_old_versions(pedido_id, keep=5)     # Limpa versões antigas
"""

import os
import shutil
import json
import hashlib
from datetime import datetime
from pathlib import Path
import traceback


class VersionManager:
    """Gerenciador de versões para arquivos Master CDR."""
    
    def __init__(self, output_dir):
        """
        Inicializa gerenciador de versões.
        
        Args:
            output_dir: Caminho da pasta de saída (onde fica output/)
        """
        self.output_dir = output_dir
        self.versions_dir = os.path.join(output_dir, ".versions")
        self.manifest_path = os.path.join(self.versions_dir, "manifest.json")
        
        # Cria pasta .versions se não existir
        os.makedirs(self.versions_dir, exist_ok=True)
        
        # Inicializa manifest se não existir
        if not os.path.exists(self.manifest_path):
            self._init_manifest()
    
    def _init_manifest(self):
        """Cria arquivo manifest.json vazio."""
        manifest = {
            "created_at": datetime.now().isoformat(),
            "version_count": 0,
            "pedidos": {}
        }
        self._write_manifest(manifest)
    
    def _read_manifest(self):
        """Lê manifest.json."""
        try:
            with open(self.manifest_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {"pedidos": {}, "version_count": 0}
    
    def _write_manifest(self, manifest):
        """Escreve manifest.json."""
        try:
            with open(self.manifest_path, 'w', encoding='utf-8') as f:
                json.dump(manifest, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"[WARN] Erro ao salvar manifest: {e}")
    
    def _calculate_md5(self, filepath):
        """Calcula hash MD5 de um arquivo."""
        try:
            md5 = hashlib.md5()
            with open(filepath, 'rb') as f:
                for chunk in iter(lambda: f.read(4096), b''):
                    md5.update(chunk)
            return md5.hexdigest()
        except Exception as e:
            print(f"[ERRO] MD5 de {filepath}: {e}")
            return None
    
    def create_backup(self, master_cdr_path, pedido_id):
        """
        Cria backup versionado do arquivo Master CDR.
        
        Args:
            master_cdr_path: Caminho do arquivo O.P_XXX_Master.cdr
            pedido_id: ID do pedido (ex: "PED-001")
            
        Returns:
            dict com metadados da versão criada
        """
        try:
            if not os.path.exists(master_cdr_path):
                print(f"[ERRO] Arquivo não encontrado: {master_cdr_path}")
                return None
            
            manifest = self._read_manifest()
            
            # Inicializa entrada de pedido se não existir
            if pedido_id not in manifest["pedidos"]:
                manifest["pedidos"][pedido_id] = {
                    "first_created": datetime.now().isoformat(),
                    "versions": []
                }
            
            # Calcula número da versão
            existing_versions = manifest["pedidos"][pedido_id]["versions"]
            version_num = len(existing_versions) + 1
            version_name = f"v{version_num:03d}"
            
            # Nome do arquivo backup
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            backup_filename = f"{version_name}_{pedido_id}_{timestamp}.cdr"
            backup_path = os.path.join(self.versions_dir, backup_filename)
            
            # Copia arquivo
            shutil.copy2(master_cdr_path, backup_path)
            
            # Calcula metadados
            file_size = os.path.getsize(backup_path)
            file_md5 = self._calculate_md5(backup_path)
            
            # Registra no manifest
            version_data = {
                "version": version_name,
                "timestamp": datetime.now().isoformat(),
                "filename": backup_filename,
                "file_path": backup_path,
                "file_size_bytes": file_size,
                "md5_hash": file_md5,
                "timestamp_created": timestamp
            }
            
            manifest["pedidos"][pedido_id]["versions"].append(version_data)
            manifest["version_count"] = sum(len(p["versions"]) for p in manifest["pedidos"].values())
            
            self._write_manifest(manifest)
            
            print(f"[OK] Backup criado: {version_name} ({file_size/1024/1024:.2f}MB) → {backup_filename}")
            return version_data
            
        except Exception as e:
            print(f"[ERRO] Ao criar backup: {e}")
            traceback.print_exc()
            return None
    
    def list_versions(self, pedido_id):
        """
        Lista todas as versões de um pedido.
        
        Args:
            pedido_id: ID do pedido
            
        Returns:
            Lista de versões ou [] se não houver
        """
        manifest = self._read_manifest()
        
        if pedido_id not in manifest["pedidos"]:
            return []
        
        return manifest["pedidos"][pedido_id]["versions"]
    
    def restore_version(self, pedido_id, version_name):
        """
        Restaura uma versão anterior para a pasta de saída.
        
        Args:
            pedido_id: ID do pedido
            version_name: Nome da versão (ex: "v001")
            
        Returns:
            Caminho do arquivo restaurado ou None
        """
        try:
            manifest = self._read_manifest()
            
            if pedido_id not in manifest["pedidos"]:
                print(f"[ERRO] Pedido '{pedido_id}' não encontrado no manifest.")
                return None
            
            versions = manifest["pedidos"][pedido_id]["versions"]
            version_data = None
            
            for v in versions:
                if v["version"] == version_name:
                    version_data = v
                    break
            
            if not version_data:
                print(f"[ERRO] Versão '{version_name}' não encontrada para pedido '{pedido_id}'.")
                return None
            
            source_path = version_data["file_path"]
            
            if not os.path.exists(source_path):
                print(f"[ERRO] Arquivo de versão não existe: {source_path}")
                return None
            
            # Cria arquivo restaurado
            restore_filename = f"RESTORED_{version_name}_{pedido_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.cdr"
            restore_path = os.path.join(self.output_dir, restore_filename)
            
            shutil.copy2(source_path, restore_path)
            
            print(f"[OK] Versão {version_name} restaurada para: {restore_filename}")
            return restore_path
            
        except Exception as e:
            print(f"[ERRO] Ao restaurar versão: {e}")
            traceback.print_exc()
            return None
    
    def cleanup_old_versions(self, pedido_id, keep=5):
        """
        Remove versões antigas, mantendo apenas as N últimas.
        
        Args:
            pedido_id: ID do pedido
            keep: Número de versões a manter (default: 5)
            
        Returns:
            Número de versões deletadas
        """
        try:
            manifest = self._read_manifest()
            
            if pedido_id not in manifest["pedidos"]:
                print(f"[WARN] Pedido '{pedido_id}' não encontrado.")
                return 0
            
            versions = manifest["pedidos"][pedido_id]["versions"]
            
            if len(versions) <= keep:
                print(f"[INFO] Apenas {len(versions)} versões. Nada a limpar (keep={keep}).")
                return 0
            
            # Ordena por timestamp (mais antigas primeiro)
            versions_sorted = sorted(versions, key=lambda x: x["timestamp"])
            
            # Remove as mais antigas
            versions_to_delete = versions_sorted[:-keep]
            deleted_count = 0
            
            for v in versions_to_delete:
                try:
                    filepath = v["file_path"]
                    if os.path.exists(filepath):
                        os.remove(filepath)
                        deleted_count += 1
                        print(f"   Deletado: {v['filename']}")
                except Exception as e:
                    print(f"   [WARN] Erro ao deletar {v['filename']}: {e}")
            
            # Atualiza manifest
            manifest["pedidos"][pedido_id]["versions"] = versions[-keep:]
            manifest["version_count"] = sum(len(p["versions"]) for p in manifest["pedidos"].values())
            self._write_manifest(manifest)
            
            print(f"[OK] Limpeza concluída: {deleted_count} versões deletadas. Mantidas: {keep}")
            return deleted_count
            
        except Exception as e:
            print(f"[ERRO] Ao limpar versões: {e}")
            traceback.print_exc()
            return 0
    
    def export_report(self, output_file="version_report.txt"):
        """
        Exporta relatório em TXT com histórico completo.
        
        Args:
            output_file: Caminho do arquivo de saída
            
        Returns:
            Caminho do arquivo gerado
        """
        try:
            manifest = self._read_manifest()
            
            report_lines = [
                "=" * 80,
                "RELATÓRIO DE VERSIONAMENTO — MALHARIA ESPORTIVA 5.0",
                "=" * 80,
                f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
                f"Total de Versões: {manifest['version_count']}",
                ""
            ]
            
            for pedido_id, pedido_data in manifest["pedidos"].items():
                report_lines.append(f"\n📦 PEDIDO: {pedido_id}")
                report_lines.append(f"   Primeira Criação: {pedido_data['first_created']}")
                report_lines.append(f"   Total de Versões: {len(pedido_data['versions'])}")
                report_lines.append("   " + "-" * 76)
                
                for v in pedido_data["versions"]:
                    size_mb = v["file_size_bytes"] / 1024 / 1024
                    report_lines.append(f"   {v['version']} | {v['timestamp']}")
                    report_lines.append(f"           Arquivo: {v['filename']}")
                    report_lines.append(f"           Tamanho: {size_mb:.2f} MB")
                    report_lines.append(f"           MD5: {v['md5_hash']}")
                    report_lines.append("")
            
            report_lines.append("=" * 80)
            
            report_content = "\n".join(report_lines)
            
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(report_content)
            
            print(f"[OK] Relatório exportado: {output_file}")
            return output_file
            
        except Exception as e:
            print(f"[ERRO] Ao exportar relatório: {e}")
            traceback.print_exc()
            return None
