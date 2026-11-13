# Gerador de Propostas — Solar In Focco (janela simples)
import datetime as dt, os, queue, re, sys, threading, traceback
from pathlib import Path

APP = Path(__file__).resolve().parent
sys.path.insert(0, str(APP))

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

AZUL = "#0B2A5B"; VERDE = "#1F8B3B"; LARANJA = "#E8730C"; CINZA = "#5F6368"; FUNDO = "#FFFFFF"

def pasta_documentos():
    if os.name == "nt":
        try:
            import ctypes
            buf = ctypes.create_unicode_buffer(260)
            ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buf)
            if buf.value:
                return Path(buf.value)
        except Exception:
            pass
    return Path.home() / "Documents"

SAIDA_BASE = pasta_documentos() / "Propostas Solar In Focco"

def abrir(caminho):
    try:
        if os.name == "nt":
            os.startfile(str(caminho))
        elif sys.platform == "darwin":
            os.system(f'open "{caminho}"')
        else:
            os.system(f'xdg-open "{caminho}"')
    except Exception as e:
        messagebox.showerror("Não consegui abrir", str(e))

class App:
    def __init__(self, root):
        self.root = root
        self.conta = None
        self.orcamento = None
        self.resultado = None
        self.fila = queue.Queue()
        root.title("Gerador de Propostas — Solar In Focco")
        root.configure(bg=FUNDO)
        root.geometry("680x720")
        root.minsize(620, 640)
        try:
            root.iconbitmap(str(APP / "icone.ico"))
        except Exception:
            pass

        st = ttk.Style()
        try:
            st.theme_use("clam")
        except Exception:
            pass
        st.configure("TProgressbar", thickness=10)

        topo = tk.Frame(root, bg=FUNDO); topo.pack(pady=(14, 4))
        try:
            self.logo = tk.PhotoImage(file=str(APP / "logo.png")).subsample(2, 2)
            tk.Label(topo, image=self.logo, bg=FUNDO).pack()
        except Exception:
            tk.Label(topo, text="Solar In Focco", font=("Segoe UI", 22, "bold"), fg=AZUL, bg=FUNDO).pack()
        tk.Label(topo, text="Gerador de Propostas", font=("Segoe UI", 16, "bold"), fg=AZUL, bg=FUNDO).pack()

        corpo = tk.Frame(root, bg=FUNDO); corpo.pack(fill="x", padx=36, pady=8)
        self.b1, self.l1 = self._passo(corpo, "1", "Escolher a CONTA DE ENERGIA (PDF)", self.escolher_conta)
        self.b2, self.l2 = self._passo(corpo, "2", "Escolher o ORÇAMENTO DA DISTRIBUIDORA (PDF)", self.escolher_orcamento)

        self.bgerar = tk.Button(corpo, text="3   GERAR PROPOSTA", font=("Segoe UI", 15, "bold"),
                                bg="#BDBDBD", fg="white", activebackground=VERDE, activeforeground="white",
                                relief="flat", cursor="hand2", state="disabled", height=2, command=self.gerar)
        self.bgerar.pack(fill="x", pady=(14, 6))

        self.status = tk.Label(corpo, text="Escolha os dois arquivos para começar.", font=("Segoe UI", 11),
                               fg=CINZA, bg=FUNDO)
        self.status.pack(pady=(4, 2))
        self.barra = ttk.Progressbar(corpo, mode="indeterminate")

        self.caixa = tk.Frame(root, bg=FUNDO)
        self.caixa.pack(fill="both", expand=True, padx=36, pady=(4, 12))

    def _passo(self, pai, num, texto, cmd):
        f = tk.Frame(pai, bg=FUNDO); f.pack(fill="x", pady=6)
        b = tk.Button(f, text=f"{num}   {texto}", font=("Segoe UI", 12, "bold"), bg=AZUL, fg="white",
                      activebackground="#16407F", activeforeground="white", relief="flat", cursor="hand2",
                      anchor="w", padx=14, pady=10, command=cmd)
        b.pack(fill="x")
        l = tk.Label(f, text="Nenhum arquivo escolhido", font=("Segoe UI", 10), fg=CINZA, bg=FUNDO, anchor="w")
        l.pack(fill="x", padx=4, pady=(3, 0))
        return b, l

    def _escolher(self, titulo):
        return filedialog.askopenfilename(title=titulo, filetypes=[("Arquivo PDF", "*.pdf"), ("Todos os arquivos", "*.*")])

    def escolher_conta(self):
        f = self._escolher("Escolha a conta de energia (PDF)")
        if f:
            self.conta = f
            self.l1.config(text="✔  " + Path(f).name, fg=VERDE)
        self._atualiza()

    def escolher_orcamento(self):
        f = self._escolher("Escolha o orçamento da distribuidora (PDF)")
        if f:
            self.orcamento = f
            self.l2.config(text="✔  " + Path(f).name, fg=VERDE)
        self._atualiza()

    def _atualiza(self):
        ok = bool(self.conta and self.orcamento)
        self.bgerar.config(state="normal" if ok else "disabled", bg=VERDE if ok else "#BDBDBD")
        if ok:
            self.status.config(text="Tudo pronto. Clique em GERAR PROPOSTA.", fg=CINZA)

    def gerar(self):
        for w in self.caixa.winfo_children():
            w.destroy()
        for b in (self.b1, self.b2, self.bgerar):
            b.config(state="disabled")
        self.barra.pack(fill="x", pady=6)
        self.barra.start(12)
        threading.Thread(target=self._trabalho, daemon=True).start()
        self.root.after(150, self._ouvir)

    def _trabalho(self):
        try:
            import gerar_proposta as gp
            import shutil
            tmp = SAIDA_BASE / "_gerando"
            shutil.rmtree(tmp, ignore_errors=True)
            res = gp.gerar(self.conta, self.orcamento, tmp, progresso=lambda m: self.fila.put(("msg", m)))
            cliente = re.sub(r'[<>:"/\\|?*]', "", res["calculo"]["cliente"] or "Cliente").strip()
            destino = SAIDA_BASE / f"{cliente} - {dt.date.today().strftime('%d-%m-%Y')}"
            n = 2
            while destino.exists():
                destino = SAIDA_BASE / f"{cliente} - {dt.date.today().strftime('%d-%m-%Y')} ({n})"; n += 1
            tmp.rename(destino)
            for k in ("docx", "pdf", "xlsx"):
                if res.get(k):
                    res[k] = destino / Path(res[k]).name
            res["pasta"] = destino
            self.fila.put(("ok", res))
        except Exception as e:
            try:
                (APP / "erros.log").write_text(traceback.format_exc(), encoding="utf-8")
            except Exception:
                pass
            self.fila.put(("erro", str(e)))

    def _ouvir(self):
        try:
            while True:
                tipo, dado = self.fila.get_nowait()
                if tipo == "msg":
                    self.status.config(text=dado, fg=CINZA)
                elif tipo == "ok":
                    self._fim(); self._mostrar(dado); return
                elif tipo == "erro":
                    self._fim()
                    self.status.config(text="Não foi possível gerar a proposta.", fg="#C62828")
                    messagebox.showerror("Ops, algo deu errado",
                                         "Não foi possível gerar a proposta.\n\n" + dado +
                                         "\n\nConfira se escolheu a conta e o orçamento certos. "
                                         "Se continuar, envie o arquivo 'erros.log' da pasta do programa para o suporte.")
                    return
        except queue.Empty:
            pass
        self.root.after(150, self._ouvir)

    def _fim(self):
        self.barra.stop(); self.barra.pack_forget()
        for b in (self.b1, self.b2):
            b.config(state="normal")
        self._atualiza()

    def _mostrar(self, res):
        import gerar_proposta as gp
        r = res["calculo"]
        self.status.config(text="✔  Proposta pronta!", fg=VERDE, font=("Segoe UI", 13, "bold"))
        info = tk.Frame(self.caixa, bg="#F2F6FB", padx=14, pady=10); info.pack(fill="x")
        linhas = [("Cliente", r["cliente"]),
                  ("Sistema", f"{gp.br(r['kwp'])} kWp · {r['qtd_paineis']} painéis {r['modulo_w']} W"),
                  ("Valor da proposta", f"R$ {gp.br(r['investimento'])}"),
                  ("Retorno do investimento", f"{r['ano_payback']}º ano")]
        for i, (k, v) in enumerate(linhas):
            tk.Label(info, text=k + ":", font=("Segoe UI", 10), fg=CINZA, bg="#F2F6FB").grid(row=i, column=0, sticky="w")
            tk.Label(info, text=v, font=("Segoe UI", 10, "bold"), fg=AZUL, bg="#F2F6FB").grid(row=i, column=1, sticky="w", padx=8)

        avisos = list(res["alertas"]) + [f"Campo não encontrado no modelo: {c}" for c in res["nao_encontrados"]]
        if avisos:
            tk.Label(self.caixa, text="⚠  Confira antes de enviar ao cliente:", font=("Segoe UI", 10, "bold"),
                     fg=LARANJA, bg=FUNDO, anchor="w").pack(fill="x", pady=(10, 2))
            for a in avisos:
                tk.Label(self.caixa, text="•  " + a, font=("Segoe UI", 9), fg="#6D4C00", bg=FUNDO,
                         wraplength=580, justify="left", anchor="w").pack(fill="x")

        bts = tk.Frame(self.caixa, bg=FUNDO); bts.pack(fill="x", pady=(12, 0))
        def bt(txt, alvo, cor):
            tk.Button(bts, text=txt, font=("Segoe UI", 11, "bold"), bg=cor, fg="white", relief="flat",
                      cursor="hand2", padx=10, pady=6, command=lambda: abrir(alvo)).pack(side="left", expand=True, fill="x", padx=4)
        bt("Abrir proposta (Word)", res["docx"], AZUL)
        if res.get("pdf") and Path(res["pdf"]).exists():
            bt("Abrir PDF", res["pdf"], "#C62828")
        bt("Abrir pasta", res["pasta"], CINZA)

if __name__ == "__main__":
    try:
        root = tk.Tk()
        App(root)
        root.mainloop()
    except Exception:
        erro = traceback.format_exc()
        try:
            (APP / "erros.log").write_text(erro, encoding="utf-8")
        except Exception:
            pass
        try:
            messagebox.showerror("Erro ao abrir o Gerador", erro)
        except Exception:
            pass
        raise
