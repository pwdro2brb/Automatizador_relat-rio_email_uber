import re
import sys
from pathlib import Path
 
import win32com.client as win32

# =========================
# CONFIGURAÇÕES
# =========================
ASSUNTO = "Referente a utilização de uber no centro de custo período março - abril"

CORPO = (
    "Segue em anexo as utilizações do Uber coorporativo referente aos meses março e abril de 2026, "
    "solicitados nos centros de custos sob sua responsabilidade. Caso estiver de acordo, não é necessário responder esse e-mail."
)

# CC opcional (deixe vazio). Depois você pode preencher:
# CC_OPCIONAL = "email1@mrv.com.br;email2@mrv.com.br"
CC_OPCIONAL = "vanessa.brodrigues@mrv.com.br" 

# extensões consideradas como anexo
EXTENSOES = {".xlsx", ".xlsm", ".xls"}

# =========================
# FUNÇÕES
# =========================
def achar_pasta_unica(base_dir: Path) -> Path:
    pastas = [p for p in base_dir.iterdir() if p.is_dir()]
    if len(pastas) != 1:
        raise RuntimeError(
            f"Era esperado existir APENAS 1 pasta aqui. Encontrei: {[p.name for p in pastas]}"
        )
    return pastas[0]

def limpar_nome_arquivo(nome: str) -> str:
    """Remove extensão e espaços extras; isso vira o DESTINATÁRIO."""
    nome = Path(nome).stem
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome

# =========================
# MAIN
# =========================
def main():
    base_dir = Path(__file__).resolve().parent
    pasta_mes = achar_pasta_unica(base_dir)

    arquivos = sorted(
        [f for f in pasta_mes.iterdir() if f.is_file() and f.suffix.lower() in EXTENSOES]
    )

    if not arquivos:
        print(f"❌ Nenhum arquivo Excel encontrado dentro de: {pasta_mes}")
        sys.exit(1)

    outlook = win32.Dispatch("Outlook.Application")

    falhas = []

    for f in arquivos:
        destinatario = limpar_nome_arquivo(f.name)

        mail = outlook.CreateItem(0)  # 0 = MailItem
        mail.To = destinatario
        mail.CC = CC_OPCIONAL
        mail.Subject = ASSUNTO

        # anexo
        mail.Attachments.Add(str(f))

        # tenta resolver destinatários (bom para validar se existe na GAL/contatos)
        recip = mail.Recipients
        recip.ResolveAll()

        resolved = True
        for i in range(1, recip.Count + 1):
            if not recip.Item(i).Resolved:
                resolved = False

        if not resolved:
            falhas.append(destinatario)


        # Força HTML (necessário para assinatura com imagem)
        mail.BodyFormat = 2  # 2 = olFormatHTML

        # Cria o e-mail e deixa o Outlook inserir a assinatura
        mail.Display(False)  # abre em modo não-modal (não trava)

        # Agora o HTMLBody já deve conter a assinatura padrão
        assinatura_html = mail.HTMLBody

        # Coloca seu texto acima da assinatura
        mail.HTMLBody = f"""
        <html>
        <body>
            <p>{CORPO}</p>
            <br>
            {assinatura_html}
        </body>
        </html>
        """

        # Salva como rascunho
        mail.Save()

        # (Opcional) Fecha a janela que abriu, sem perguntar nada
        mail.Close(0)  # 0 = olDiscard (descarta alterações da janela, mas já salvou)


        print(f"✅ Rascunho criado: {destinatario} | Anexo: {f.name}")

    if falhas:
        print("\n⚠️ ATENÇÃO: alguns destinatários NÃO foram resolvidos pelo Outlook (nome não encontrado):")
        for n in falhas:
            print(" -", n)
        print("\nDica: nesses casos você pode renomear o arquivo para o e-mail.Body, ou usar um mapeamento Nome→E-mail.")
    else:
        print("\n✅ Todos os destinatários foram resolvidos com sucesso!")

    print(f"\n📁 Pasta processada: {pasta_mes.name}")
    print("Fim.")

if __name__ == "__main__":
    main()
