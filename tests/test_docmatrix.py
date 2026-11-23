from __future__ import annotations
import json
from pathlib import Path
import sys
import tempfile
import unittest
import zipfile

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from scripts.docmatrix import prepare,finalize
from scripts.extract import extract_document
from scripts.normalize import normalize_value,excel_safe
from scripts.verify import verify_outputs


class DocMatrixTest(unittest.TestCase):
    def test_normalize_preserves_identifiers(self):
        self.assertEqual(normalize_value('00123'),('00123','texto'))
        self.assertEqual(normalize_value('23/11/2026'),('2026-11-23','data'))
        self.assertEqual(normalize_value('R$ 1.250,30'),(1250.3,'numero'))
        self.assertEqual(normalize_value('9999999999999999'),('9999999999999999','identificador'))
        self.assertEqual(excel_safe('=SUM(A1:A2)'),"'=SUM(A1:A2)")

    def test_text_and_csv(self):
        a=extract_document(ROOT/'examples/exemplo_a.txt','doc_001')
        self.assertEqual(a['status'],'processado')
        self.assertTrue(any(r['campo']=='Orçamento total' for r in a['registros']))
        self.assertTrue(any(e['data']=='2026-11-12' for e in a['eventos']))
        tab=extract_document(ROOT/'examples/exemplo_tabela.csv','doc_002')
        self.assertEqual(len(tab['tabelas']),1)
        self.assertEqual(len(tab['tabelas'][0]['linhas']),3)
        self.assertEqual(tab['tabelas'][0]['colunas'][0],'produto')

    def test_two_phase_and_artifact(self):
        with tempfile.TemporaryDirectory() as t:
            base,ctx=prepare([str(ROOT/'examples/exemplo_a.txt'),str(ROOT/'examples/exemplo_b.txt')],t,'comparacao','Orçamento total')
            self.assertTrue(base.exists()); self.assertTrue(ctx.exists())
            payload=json.loads(base.read_text(encoding='utf-8'))
            comp={x['campo']:x for x in payload['comparativo']}
            self.assertEqual(comp['Orçamento total']['status'],'diferente')
            self.assertEqual(comp['Projeto']['status'],'igual')
            annotation={'analise':{'resumo':'Diferença de orçamento e prazo, baseada nos documentos.'}}
            annot=Path(t)/'anotacoes.json';annot.write_text(json.dumps(annotation,ensure_ascii=False),encoding='utf-8')
            result=finalize(str(base),t,str(annot))
            check=verify_outputs(result['json'],result['excel'])
            self.assertEqual(check['status'],'aprovado')
            self.assertIn('Comparativo',check['abas'])
            self.assertIn('Síntese',check['abas'])
            complete=json.loads(Path(result['json']).read_text(encoding='utf-8'))
            self.assertEqual(complete['execucao']['interpretacao_ia'],'fornecida_pelo_agente')
            read_back=extract_document(result['excel'],'doc_007')
            self.assertEqual(read_back['status'],'processado')
            self.assertGreaterEqual(len(read_back['tabelas']),4)

    def test_json_xml_and_docx_fields(self):
        with tempfile.TemporaryDirectory() as t:
            base=Path(t)
            j=base/'data.json'
            j.write_text('{"codigo":"00123","total":"1.250,00"}',encoding='utf-8')
            out=extract_document(j,'doc_001')
            self.assertEqual(out['status'],'processado')
            self.assertTrue(any(r['campo']=='$.codigo' for r in out['registros']))
            x=base/'data.xml'; x.write_text('<root><total>1250</total></root>',encoding='utf-8')
            out=extract_document(x,'doc_002')
            self.assertTrue(any(r['campo']=='/root/total' for r in out['registros']))
            d=base/'data.docx'
            xml=('<?xml version="1.0"?>'
                 '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                 '<w:body><w:p><w:r><w:t>Assunto: Teste</w:t></w:r></w:p>'
                 '</w:body></w:document>')
            with zipfile.ZipFile(d,'w') as z: z.writestr('word/document.xml',xml)
            out=extract_document(d,'doc_003')
            self.assertEqual(out['status'],'processado')
            self.assertTrue(any(r['campo']=='Assunto' for r in out['registros']))

    def test_pdf_when_pymupdf_installed(self):
        try:
            import fitz
        except ImportError:
            self.skipTest('PyMuPDF não disponível')
        with tempfile.TemporaryDirectory() as t:
            p=Path(t)/'sample.pdf'
            doc=fitz.open()
            page=doc.new_page(); page.insert_text((72,72),'Field: Value')
            doc.save(p);doc.close()
            out=extract_document(p,'doc_001')
            self.assertEqual(out['status'],'processado')
            self.assertEqual(out['paginas'],1)
            self.assertTrue(any(r['campo']=='Field' for r in out['registros']))

    def test_xlsxwriter_backend_and_injection(self):
        from scripts.export_excel import export_excel
        import xml.etree.ElementTree as ET
        with tempfile.TemporaryDirectory() as t:
            base, _ = prepare([str(ROOT / 'examples/exemplo_a.txt'),
                               str(ROOT / 'examples/exemplo_b.txt')], t, 'comparacao')
            payload = json.loads(base.read_text(encoding='utf-8'))
            payload['documentos'][0]['registros'].append({
                'id': 'doc_001:rEXTRA', 'campo': 'Campo perigoso',
                'valor_original': '=HYPERLINK("https://example.com")',
                'valor_normalizado': '=2+2', 'tipo': 'texto', 'origem': 'teste',
                'pagina': 1, 'localizacao': 'teste', 'bloco_id': 'doc_001:b000001',
            })
            output = Path(t) / 'docmatrix.xlsx'
            sheets, backend = export_excel(payload, output, backend='xlsxwriter')
            self.assertEqual(backend, 'xlsxwriter')
            self.assertIn('Comparativo', sheets)
            with zipfile.ZipFile(output) as z:
                main = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
                for sheetname in [n for n in z.namelist() if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')]:
                    xml = ET.fromstring(z.read(sheetname))
                    self.assertEqual(xml.findall(f'.//{main}f'), [], 'Texto não confiável virou fórmula')
            rec = extract_document(output, 'doc_out')
            self.assertEqual(rec['status'], 'processado')
            self.assertTrue(any(tbl['localizacao'] == 'aba Comparativo' for tbl in rec['tabelas']))

    def test_xlsxwriter_many_rows(self):
        from scripts.export_excel import ExcelWriter
        with tempfile.TemporaryDirectory() as t:
            file = Path(t) / 'big.xlsx'
            writer = ExcelWriter(file, backend='xlsxwriter')
            writer.write_tab('Dados', ['index', 'texto'], [[i, f'linha {i}'] for i in range(2500)])
            writer.close()
            self.assertEqual(writer.counts['Dados'], 2500)
            doc = extract_document(file, 'doc_out')
            self.assertEqual(len(doc['tabelas'][0]['linhas']), 2500)

    def test_installer_project_paths(self):
        from install import install
        with tempfile.TemporaryDirectory() as t:
            target_claude = install('claude', 'project', t)
            self.assertTrue((target_claude / 'SKILL.md').is_file())
            self.assertTrue((target_claude / 'requirements.txt').is_file())
            self.assertIn('.claude/skills/docmatrix', target_claude.as_posix())
            with self.assertRaises(FileExistsError):
                install('claude', 'project', t)
            self.assertEqual(install('claude', 'project', t, overwrite=True), target_claude)
            target_codex = install('codex', 'project', t)
            self.assertTrue((target_codex / 'SKILL.md').exists())
            self.assertIn('.agents/skills/docmatrix', target_codex.as_posix())

    def test_xml_blocks_doctype(self):
        with tempfile.TemporaryDirectory() as t:
            p=Path(t)/'bad.xml'; p.write_text('<!DOCTYPE root [<!ENTITY ext SYSTEM "file:///etc/passwd">]><root>&ext;</root>')
            result=extract_document(p,'doc_001')
            self.assertEqual(result['status'],'erro')
            self.assertTrue(any('DTD' in w for w in result['avisos']))

if __name__=='__main__': unittest.main()
