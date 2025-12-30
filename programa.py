import pyautogui as p
import time
import random
import PySimpleGUI as sg
import pandas as pd
from openpyxl import load_workbook
import os
import pyexcel as pe
import json


class ConfiguracoesUsuario:
    ARQUIVO_CONFIG = "configuracoes_usuarios.json"

    CONFIG_PADRAO = {
        "valor_max_venda": "1000",
        "margem_venda": "30",
        "porcentagem_estoque": "50",
        "estoque_minimo": "1",
        "tempo_espera_inicial": "5",
        "tempo_espera": "0.5",
        "confirmar_venda": True,
        "usar_clique_pdv": True,
        "quantidade_max_item": "200",
        "quantidade_max_por_vez": "99",
        "x_codigo": None,
        "y_codigo": None,
        "x_dinheiro": None,
        "y_dinheiro": None,
        "x_finalizar": None,
        "y_finalizar": None,
        "x_fechar": None,
        "y_fechar": None,
        "codigos_exclusao": []
    }

    @staticmethod
    def carregar():
        if os.path.exists(ConfiguracoesUsuario.ARQUIVO_CONFIG):
            try:
                with open(ConfiguracoesUsuario.ARQUIVO_CONFIG, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return {}
        return {}

    @staticmethod
    def salvar(dados):
        with open(ConfiguracoesUsuario.ARQUIVO_CONFIG, 'w', encoding='utf-8') as f:
            json.dump(dados, f, indent=2, ensure_ascii=False)

    @staticmethod
    def obter_config(usuario):
        dados = ConfiguracoesUsuario.carregar()
        if usuario in dados:
            config = ConfiguracoesUsuario.CONFIG_PADRAO.copy()
            config.update(dados[usuario])
            return config
        return ConfiguracoesUsuario.CONFIG_PADRAO.copy()

    @staticmethod
    def salvar_config(usuario, config):
        dados = ConfiguracoesUsuario.carregar()
        dados[usuario] = config
        ConfiguracoesUsuario.salvar(dados)

    @staticmethod
    def atualizar_config(usuario, **kwargs):
        config = ConfiguracoesUsuario.obter_config(usuario)
        config.update(kwargs)
        ConfiguracoesUsuario.salvar_config(usuario, config)

    @staticmethod
    def obter_lista_usuarios():
        dados = ConfiguracoesUsuario.carregar()
        return sorted(list(dados.keys()))

    @staticmethod
    def criar_usuario(usuario):
        dados = ConfiguracoesUsuario.carregar()
        if usuario not in dados:
            dados[usuario] = ConfiguracoesUsuario.CONFIG_PADRAO.copy()
            ConfiguracoesUsuario.salvar(dados)
            return True
        return False

    @staticmethod
    def obter_codigos_exclusao(usuario):
        config = ConfiguracoesUsuario.obter_config(usuario)
        return set(config.get('codigos_exclusao', []))

    @staticmethod
    def adicionar_codigo_exclusao(usuario, codigo):
        config = ConfiguracoesUsuario.obter_config(usuario)
        codigos = config.get('codigos_exclusao', [])
        codigo_str = str(codigo).strip()
        if codigo_str and codigo_str not in codigos:
            codigos.append(codigo_str)
            config['codigos_exclusao'] = codigos
            ConfiguracoesUsuario.salvar_config(usuario, config)
            return True
        return False

    @staticmethod
    def remover_codigo_exclusao(usuario, codigo):
        config = ConfiguracoesUsuario.obter_config(usuario)
        codigos = config.get('codigos_exclusao', [])
        codigo_str = str(codigo).strip()
        if codigo_str in codigos:
            codigos.remove(codigo_str)
            config['codigos_exclusao'] = codigos
            ConfiguracoesUsuario.salvar_config(usuario, config)
            return True
        return False

    @staticmethod
    def obter_lista_exclusao(usuario):
        config = ConfiguracoesUsuario.obter_config(usuario)
        return config.get('codigos_exclusao', [])


class AutoSoftcom:
    def __init__(self, path, tempo_espera_inicial, tempo_espera, valor_max_venda, x_codigo, y_codigo,
                 x_dinheiro, y_dinheiro, x_finalizar, y_finalizar, x_fechar, y_fechar, porcentagem_estoque, estoque_minimo,
                 margem_venda, confirmar_venda, usar_clique_pdv=True, usuario=None, quantidade_max_item=200, quantidade_max_por_vez=99):
        time.sleep(float(tempo_espera_inicial))

        self.tempo_espera = float(tempo_espera)
        self.valor_max_venda = float(valor_max_venda)
        self.x_codigo = x_codigo
        self.y_codigo = y_codigo
        self.x_dinheiro = x_dinheiro
        self.y_dinheiro = y_dinheiro
        self.x_finalizar = x_finalizar
        self.y_finalizar = y_finalizar
        self.x_fechar = x_fechar
        self.y_fechar = y_fechar
        self.porcentagem_estoque = float(porcentagem_estoque) / 100
        self.estoque_minimo = int(estoque_minimo)
        self.margem_venda = float(margem_venda) / 100
        self.confirmar_venda = confirmar_venda
        self.usar_clique_pdv = usar_clique_pdv
        self.usuario = usuario
        self.quantidade_max_item = int(quantidade_max_item)
        self.quantidade_max_por_vez = int(quantidade_max_por_vez)
        self.codigos_exclusao = set()

        if usuario:
            self.codigos_exclusao = ConfiguracoesUsuario.obter_codigos_exclusao(
                usuario)
            if self.codigos_exclusao:
                print(
                    f"Lista de exclusão carregada para '{usuario}': {len(self.codigos_exclusao)} códigos")

        self.path = path
        self.valor_total_vendido = 0
        self.contador_vendas = 0

        self.carregar_planilha()

    def carregar_planilha(self):
        try:
            if self.path.endswith('.xls'):
                try:
                    records = pe.get_records(file_name=self.path)
                    self.df = pd.DataFrame(list(records))
                except:
                    self.df = pd.read_excel(self.path, engine='xlrd')
            else:
                self.df = pd.read_excel(self.path, engine='openpyxl')
        except Exception as e:
            try:
                self.df = pd.read_excel(self.path)
            except:
                raise ValueError(f"Erro ao ler arquivo: {e}")

        self.df.columns = self.df.columns.str.strip()

        print(f"Colunas encontradas: {list(self.df.columns)}")

        if 'Quantidade vendida' not in self.df.columns:
            print("Coluna 'Quantidade vendida' não encontrada, criando...")
            self.df['Quantidade vendida'] = 0
        else:
            print("Coluna 'Quantidade vendida' encontrada")

        self.df['Quantidade vendida'] = pd.to_numeric(
            self.df['Quantidade vendida'], errors='coerce').fillna(0)

        print(
            f"Valores iniciais de 'Quantidade vendida': {self.df['Quantidade vendida'].head(10).tolist()}")
        self.df['Estoque'] = pd.to_numeric(
            self.df['Estoque'], errors='coerce').fillna(0)
        self.df['PrecoUnitario'] = pd.to_numeric(
            self.df['PrecoUnitario'], errors='coerce').fillna(0)
        self.df['Codigo'] = self.df['Codigo'].astype(str)

        self.is_xls = self.path.endswith('.xls')
        self.filtrar_produtos_disponiveis()

    def filtrar_produtos_disponiveis(self):
        estoque_disponivel = self.df['Estoque'] - self.df['Quantidade vendida']
        self.df['Estoque disponivel'] = estoque_disponivel
        mask = (estoque_disponivel > self.estoque_minimo) & (
            self.df['Quantidade vendida'] == 0)

        if self.codigos_exclusao:
            produtos_na_exclusao = self.df[self.df['Codigo'].astype(
                str).isin(self.codigos_exclusao)]
            produtos_excluidos_count = len(produtos_na_exclusao)
            if produtos_excluidos_count > 0:
                codigos_excluidos = produtos_na_exclusao['Codigo'].astype(
                    str).tolist()
                print(
                    f"\n🚫 Lista de exclusão ativa: {produtos_excluidos_count} produtos serão ignorados")
                print(
                    f"   Códigos excluídos: {', '.join(codigos_excluidos[:10])}{'...' if len(codigos_excluidos) > 10 else ''}")
            mask = mask & (~self.df['Codigo'].astype(
                str).isin(self.codigos_exclusao))

        self.produtos_disponiveis = self.df[mask].copy()

        if len(self.produtos_disponiveis) == 0:
            raise ValueError("Não há produtos disponíveis para venda")

    def calcular_quantidade_venda(self, estoque_disponivel, codigo_produto):
        porcentagem_pct = self.porcentagem_estoque * 100
        quantidade_exata = estoque_disponivel * self.porcentagem_estoque
        quantidade_max = int(quantidade_exata)

        print(f"\n--- Produto {codigo_produto} ---")
        print(f"Estoque disponível: {estoque_disponivel} unidades")
        print(f"Porcentagem configurada: {porcentagem_pct:.1f}%")
        print(
            f"Quantidade calculada pela porcentagem: {quantidade_exata:.2f} → {quantidade_max} unidades")

        if quantidade_max < 1:
            print(f"❌ Quantidade calculada < 1 unidade, produto ignorado")
            return 0

        quantidade_escolhida = quantidade_max

        print(
            f"✅ Quantidade final escolhida: {quantidade_escolhida} unidades ({porcentagem_pct:.1f}% do estoque)")

        return quantidade_escolhida

    def calcular_preco_venda(self, preco_custo, quantidade):
        preco_unitario_venda = preco_custo * (1 + self.margem_venda)
        return preco_unitario_venda * quantidade

    def fracionar_quantidade(self, quantidade_total):
        quantidade_total = min(quantidade_total, self.quantidade_max_item)

        if quantidade_total >= 100:
            resposta = sg.popup_yes_no(
                f"Quantidade a inserir: {quantidade_total} unidades\n\nDeseja continuar?",
                title="Confirmação de Quantidade",
                keep_on_top=True
            )
            if resposta != 'Yes':
                return []

        if quantidade_total <= self.quantidade_max_por_vez:
            return [quantidade_total]

        fracoes = []
        quantidade_restante = quantidade_total

        while quantidade_restante > 0:
            quantidade_lote = min(quantidade_restante,
                                  self.quantidade_max_por_vez)
            fracoes.append(quantidade_lote)
            quantidade_restante -= quantidade_lote

        return fracoes

    def processar_venda(self):
        produtos_vendidos = []
        valor_venda_atual = 0

        for idx, produto in self.produtos_disponiveis.iterrows():
            estoque_disponivel = produto['Estoque disponivel']
            codigo_produto = produto['Codigo']

            if estoque_disponivel <= self.estoque_minimo:
                print(
                    f"\n⚠️ Produto {codigo_produto}: Estoque ({estoque_disponivel}) <= estoque mínimo ({self.estoque_minimo}), ignorado")
                continue

            quantidade = self.calcular_quantidade_venda(
                estoque_disponivel, codigo_produto)

            if quantidade == 0:
                continue

            quantidade = min(quantidade, self.quantidade_max_item)

            preco_unitario_venda = produto['PrecoUnitario'] * \
                (1 + self.margem_venda)
            preco_venda_item_total = self.calcular_preco_venda(
                produto['PrecoUnitario'], quantidade)

            print(f"Preço unitário (custo): R$ {produto['PrecoUnitario']:.2f}")
            print(
                f"Preço unitário (venda com {self.margem_venda*100:.1f}% margem): R$ {preco_unitario_venda:.2f}")
            print(
                f"Valor total do item ({quantidade} unidades): R$ {preco_venda_item_total:.2f}")

            if valor_venda_atual + preco_venda_item_total > self.valor_max_venda:
                valor_disponivel = self.valor_max_venda - valor_venda_atual

                if valor_disponivel <= 0:
                    print(
                        f"\n⚠️ Valor máximo da venda já atingido. Finalizando venda.")
                    break

                quantidade_max_possivel_valor = int(
                    valor_disponivel / preco_unitario_venda)
                quantidade_max_possivel = min(
                    quantidade_max_possivel_valor, estoque_disponivel, self.quantidade_max_item)

                if quantidade_max_possivel < 1:
                    print(
                        f"\n⚠️ Produto {codigo_produto} não adicionado: mesmo 1 unidade ultrapassaria o valor máximo")
                    print(
                        f"   Valor atual da venda: R$ {valor_venda_atual:.2f}")
                    print(f"   Valor disponível: R$ {valor_disponivel:.2f}")
                    print(
                        f"   Valor de 1 unidade: R$ {preco_unitario_venda:.2f}")
                    print(
                        f"   Valor máximo configurado: R$ {self.valor_max_venda:.2f}")
                    break

                quantidade_original = quantidade
                quantidade = quantidade_max_possivel
                preco_venda_item_total = preco_unitario_venda * quantidade

                print(
                    f"\n⚠️ Quantidade ajustada: não foi possível adicionar {quantidade_original} unidades ({self.porcentagem_estoque*100:.1f}% do estoque)")
                print(
                    f"   Quantidade reduzida para: {quantidade} unidades (máximo possível respeitando o valor máximo)")
                print(f"   Valor ajustado: R$ {preco_venda_item_total:.2f}")
                print(
                    f"   Valor total da venda: R$ {valor_venda_atual + preco_venda_item_total:.2f} / R$ {self.valor_max_venda:.2f}")

            fracoes = self.fracionar_quantidade(quantidade)

            if not fracoes:
                print(f"⚠️ Produto {codigo_produto} cancelado pelo usuário")
                continue

            quantidade_total_inserida = 0
            preco_total_item = 0

            for i, quantidade_fracao in enumerate(fracoes):
                if self.usar_clique_pdv:
                    p.click(self.x_codigo, self.y_codigo)
                    time.sleep(self.tempo_espera)

                p.write(str(quantidade_fracao), interval=0.1)
                time.sleep(self.tempo_espera)
                p.press('*')
                time.sleep(self.tempo_espera)
                p.write(str(produto['Codigo']), interval=0.1)
                time.sleep(self.tempo_espera)
                p.press('enter')
                time.sleep(self.tempo_espera)

                quantidade_total_inserida += quantidade_fracao
                preco_fracao = preco_unitario_venda * quantidade_fracao
                preco_total_item += preco_fracao

                if len(fracoes) > 1:
                    print(
                        f"✅ Lote {i+1}/{len(fracoes)}: {quantidade_fracao} unidades do produto {codigo_produto} adicionadas")

            valor_venda_atual += preco_total_item
            produtos_vendidos.append({
                'idx': idx,
                'codigo': produto['Codigo'],
                'quantidade': quantidade_total_inserida,
                'valor': preco_total_item
            })

            if len(fracoes) > 1:
                print(
                    f"✅ Total: {quantidade_total_inserida} unidades do produto {codigo_produto} adicionadas em {len(fracoes)} lotes")
            else:
                print(
                    f"✅ {quantidade_total_inserida} unidades do produto {codigo_produto} adicionadas ao carrinho")
            print(
                f"Valor parcial da venda: R$ {valor_venda_atual:.2f} / R$ {self.valor_max_venda:.2f}")
            print("=" * 60)

        if valor_venda_atual == 0:
            return None

        p.press('f3')
        p.doubleClick(self.x_dinheiro, self.y_dinheiro)
        time.sleep(self.tempo_espera)
        p.doubleClick(self.x_finalizar, self.y_finalizar)
        time.sleep(2)

        print("")
        print("=" * 60)
        print(f"Venda concluída! Valor total: R$ {valor_venda_atual:.2f}")
        print("=" * 60)
        print("")

        self.atualizar_planilha(produtos_vendidos)

        if self.confirmar_venda:
            resposta = sg.popup_yes_no(
                "Venda concluída!\n\nProsseguir para próxima venda?",
                title="Confirmação",
                keep_on_top=True
            )
            if resposta != 'Yes':
                return 'parar'

        p.click(self.x_fechar, self.y_fechar)
        time.sleep(2)

        return valor_venda_atual

    def atualizar_planilha(self, produtos_vendidos):
        print("Atualizando planilha...")
        print(f"Produtos a atualizar: {len(produtos_vendidos)}")

        for item in produtos_vendidos:
            idx = item['idx']
            codigo = item['codigo']
            quantidade_vendida = item['quantidade']

            print(
                f"Atualizando produto {codigo} (índice {idx}): +{quantidade_vendida} unidades")

            if idx not in self.df.index:
                print(f"⚠️ Erro: Índice {idx} não encontrado no DataFrame!")
                continue

            quantidade_anterior = self.df.loc[idx, 'Quantidade vendida']
            quantidade_total_vendida = quantidade_anterior + quantidade_vendida
            self.df.loc[idx, 'Quantidade vendida'] = quantidade_total_vendida

            print(f"  Quantidade anterior: {quantidade_anterior}")
            print(f"  Quantidade vendida: {quantidade_vendida}")
            print(f"  Quantidade total: {quantidade_total_vendida}")

        df_para_salvar = self.df.copy()

        if 'Estoque disponivel' in df_para_salvar.columns:
            df_para_salvar = df_para_salvar.drop(
                columns=['Estoque disponivel'])
            print("Coluna temporária 'Estoque disponivel' removida antes de salvar")

        try:
            if self.is_xls:
                print("Salvando arquivo .xls...")
                output_path = self.path.replace('.xls', '.xlsx')
                df_para_salvar.to_excel(
                    output_path, index=False, engine='openpyxl')
                if os.path.exists(self.path):
                    os.remove(self.path)
                self.path = output_path
                self.is_xls = False
                print(f"Arquivo convertido e salvo como: {output_path}")
            else:
                print(f"Salvando arquivo .xlsx: {self.path}")
                df_para_salvar.to_excel(
                    self.path, index=False, engine='openpyxl')
        except Exception as e:
            print(f"Erro ao salvar planilha: {e}")
            print(f"Tentando método alternativo...")
            try:
                df_para_salvar.to_excel(
                    self.path, index=False, engine='openpyxl')
                print("Salvo com sucesso usando método alternativo")
            except Exception as e2:
                raise ValueError(f"Erro ao salvar planilha: {e2}")

        print("Planilha atualizada com sucesso!")
        self.filtrar_produtos_disponiveis()

    def executar(self):
        try:
            while len(self.produtos_disponiveis) > 0:
                resultado = self.processar_venda()

                if resultado == 'parar':
                    break
                elif resultado is None:
                    break

                self.valor_total_vendido += resultado
                self.contador_vendas += 1

                print(
                    f"Total vendido até agora: R$ {self.valor_total_vendido:.2f}")
                print(f"Vendas realizadas: {self.contador_vendas}")
                print("")

        except Exception as e:
            print(f"Erro durante execução: {e}")
            raise

        finally:
            p.alert(
                f"FINALIZADO\n\nTotal vendido: R$ {self.valor_total_vendido:.2f}\nVendas realizadas: {self.contador_vendas}")
            print("")
            print("!" * 60)
            print(
                f"FINALIZADO: R$ {self.valor_total_vendido:.2f}, VENDAS: {self.contador_vendas}")
            print("!" * 60)
            print("")


class WindowAuto:
    def __init__(self):
        self.x_codigo = None
        self.y_codigo = None
        self.x_dinheiro = None
        self.y_dinheiro = None
        self.x_finalizar = None
        self.y_finalizar = None
        self.x_fechar = None
        self.y_fechar = None

    def atualizar_lista_usuarios(self, janela):
        try:
            usuarios = ConfiguracoesUsuario.obter_lista_usuarios()
            janela['-USUARIO-'].update(values=usuarios,
                                       value=usuarios[0] if usuarios else "")
            if usuarios:
                self.carregar_config_usuario(janela, usuarios[0])
        except Exception as e:
            print(f"Erro ao atualizar lista de usuários: {e}")
            import traceback
            traceback.print_exc()

    def carregar_config_usuario(self, janela, usuario):
        if not usuario:
            return

        try:
            config = ConfiguracoesUsuario.obter_config(usuario)

            janela['-VALOR_MAX-'].update(config.get('valor_max_venda', '1000'))
            janela['-MARGEM-'].update(config.get('margem_venda', '30'))
            janela['-PORC_ESTOQUE-'].update(
                config.get('porcentagem_estoque', '50'))
            janela['-ESTOQUE_MIN-'].update(config.get('estoque_minimo', '1'))
            janela['-TEMPO_INICIAL-'].update(
                config.get('tempo_espera_inicial', '5'))
            janela['-TEMPO-'].update(config.get('tempo_espera', '0.5'))
            janela['-CONFIRMAR-'].update(config.get('confirmar_venda', True))
            janela['-USAR_CLIQUE_PDV-'].update(
                config.get('usar_clique_pdv', True))
            janela['-QTD_MAX_ITEM-'].update(
                config.get('quantidade_max_item', '200'))
            janela['-QTD_MAX_POR_VEZ-'].update(
                config.get('quantidade_max_por_vez', '99'))

            x_codigo = config.get('x_codigo')
            y_codigo = config.get('y_codigo')
            x_dinheiro = config.get('x_dinheiro')
            y_dinheiro = config.get('y_dinheiro')
            x_finalizar = config.get('x_finalizar')
            y_finalizar = config.get('y_finalizar')
            x_fechar = config.get('x_fechar')
            y_fechar = config.get('y_fechar')

            if x_codigo is not None and y_codigo is not None and x_codigo != "null" and y_codigo != "null":
                self.x_codigo = int(x_codigo) if isinstance(
                    x_codigo, (int, float)) else None
                self.y_codigo = int(y_codigo) if isinstance(
                    y_codigo, (int, float)) else None
                if self.x_codigo is not None and self.y_codigo is not None:
                    janela['-POS_CODIGO-'].update(
                        f"X: {self.x_codigo}, Y: {self.y_codigo}")
                else:
                    janela['-POS_CODIGO-'].update("")
            else:
                self.x_codigo = None
                self.y_codigo = None
                janela['-POS_CODIGO-'].update("")

            if x_dinheiro is not None and y_dinheiro is not None and x_dinheiro != "null" and y_dinheiro != "null":
                self.x_dinheiro = int(x_dinheiro) if isinstance(
                    x_dinheiro, (int, float)) else None
                self.y_dinheiro = int(y_dinheiro) if isinstance(
                    y_dinheiro, (int, float)) else None
                if self.x_dinheiro is not None and self.y_dinheiro is not None:
                    janela['-POS_DINHEIRO-'].update(
                        f"X: {self.x_dinheiro}, Y: {self.y_dinheiro}")
                else:
                    janela['-POS_DINHEIRO-'].update("")
            else:
                self.x_dinheiro = None
                self.y_dinheiro = None
                janela['-POS_DINHEIRO-'].update("")

            if x_finalizar is not None and y_finalizar is not None and x_finalizar != "null" and y_finalizar != "null":
                self.x_finalizar = int(x_finalizar) if isinstance(
                    x_finalizar, (int, float)) else None
                self.y_finalizar = int(y_finalizar) if isinstance(
                    y_finalizar, (int, float)) else None
                if self.x_finalizar is not None and self.y_finalizar is not None:
                    janela['-POS_FINALIZAR-'].update(
                        f"X: {self.x_finalizar}, Y: {self.y_finalizar}")
                else:
                    janela['-POS_FINALIZAR-'].update("")
            else:
                self.x_finalizar = None
                self.y_finalizar = None
                janela['-POS_FINALIZAR-'].update("")

            if x_fechar is not None and y_fechar is not None and x_fechar != "null" and y_fechar != "null":
                self.x_fechar = int(x_fechar) if isinstance(
                    x_fechar, (int, float)) else None
                self.y_fechar = int(y_fechar) if isinstance(
                    y_fechar, (int, float)) else None
                if self.x_fechar is not None and self.y_fechar is not None:
                    janela['-POS_FECHAR-'].update(
                        f"X: {self.x_fechar}, Y: {self.y_fechar}")
                else:
                    janela['-POS_FECHAR-'].update("")
            else:
                self.x_fechar = None
                self.y_fechar = None
                janela['-POS_FECHAR-'].update("")

            lista_exclusao = ConfiguracoesUsuario.obter_lista_exclusao(usuario)
            janela['-LISTA_EXCLUSAO-'].update(values=lista_exclusao)
        except Exception as e:
            print(f"Erro ao carregar configurações: {e}")

    def salvar_config_usuario(self, janela, usuario, values):
        if not usuario:
            return

        try:
            config = {
                'valor_max_venda': values.get('-VALOR_MAX-', '1000'),
                'margem_venda': values.get('-MARGEM-', '30'),
                'porcentagem_estoque': values.get('-PORC_ESTOQUE-', '50'),
                'estoque_minimo': values.get('-ESTOQUE_MIN-', '1'),
                'tempo_espera_inicial': values.get('-TEMPO_INICIAL-', '5'),
                'tempo_espera': values.get('-TEMPO-', '0.5'),
                'confirmar_venda': values.get('-CONFIRMAR-', True),
                'usar_clique_pdv': values.get('-USAR_CLIQUE_PDV-', True),
                'quantidade_max_item': values.get('-QTD_MAX_ITEM-', '200'),
                'quantidade_max_por_vez': values.get('-QTD_MAX_POR_VEZ-', '99'),
                'x_codigo': self.x_codigo,
                'y_codigo': self.y_codigo,
                'x_dinheiro': self.x_dinheiro,
                'y_dinheiro': self.y_dinheiro,
                'x_finalizar': self.x_finalizar,
                'y_finalizar': self.y_finalizar,
                'x_fechar': self.x_fechar,
                'y_fechar': self.y_fechar,
                'codigos_exclusao': ConfiguracoesUsuario.obter_lista_exclusao(usuario)
            }

            ConfiguracoesUsuario.salvar_config(usuario, config)
        except Exception as e:
            print(f"Erro ao salvar configurações: {e}")

    def atualizar_lista_exclusao(self, janela, usuario):
        if usuario:
            lista = ConfiguracoesUsuario.obter_lista_exclusao(usuario)
            janela['-LISTA_EXCLUSAO-'].update(values=lista)
        else:
            janela['-LISTA_EXCLUSAO-'].update(values=[])

    def criar_janela(self):
        col_esquerda = [
            [sg.Frame("Usuário/Comércio",
                      [
                          [sg.Text("Selecione o usuário:", size=(18, 1)),
                           sg.Combo(values=[], key="-USUARIO-", size=(25, 1), enable_events=True,
                                    readonly=False, default_value=""),
                           sg.Button("+ Novo", key="-NOVO_USUARIO-", size=(8, 1))]
                      ], expand_x=True
                      )],
            [sg.Frame("Arquivo",
                      [
                          [sg.Text('Caminho do arquivo:', size=(18, 1))],
                          [sg.InputText(key='-PATH-', size=(None, 1), enable_events=True, expand_x=True),
                           sg.FileBrowse(file_types=(("Excel", "*.xls *.xlsx"), ("Todos", "*.*")), size=(12, 1))]
                      ], expand_x=True
                      )],
            [sg.Frame("Configurações de Venda",
                      [
                          [sg.Text("Valor máximo total da venda:", size=(28, 1)),
                           sg.InputText(key="-VALOR_MAX-", size=(12, 1), default_text="1000",
                                        tooltip="Valor total máximo que a venda pode atingir",
                                        enable_events=True)],
                          [sg.Text("Margem de venda (%):", size=(28, 1)),
                           sg.InputText(key="-MARGEM-", size=(12, 1), default_text="30",
                                        tooltip="Margem aplicada sobre o preço de custo",
                                        enable_events=True)],
                          [sg.Text("Quantidade (% do estoque):", size=(28, 1)),
                           sg.InputText(key="-PORC_ESTOQUE-", size=(12, 1), default_text="50",
                                        tooltip="Porcentagem do estoque que será vendida",
                                        enable_events=True)],
                          [sg.Text("Estoque mínimo:", size=(28, 1)),
                           sg.InputText(key="-ESTOQUE_MIN-", size=(12, 1), default_text="1",
                                        tooltip="Estoque mínimo abaixo do qual não será vendido",
                                        enable_events=True)],
                          [sg.Text("Qtd. máxima por item:", size=(28, 1)),
                           sg.InputText(key="-QTD_MAX_ITEM-", size=(12, 1), default_text="200",
                                        tooltip="Quantidade máxima de um item a ser adicionado",
                                        enable_events=True)],
                          [sg.Text("Qtd. máxima por vez:", size=(28, 1)),
                           sg.InputText(key="-QTD_MAX_POR_VEZ-", size=(12, 1), default_text="99",
                                        tooltip="Quantidade máxima para colocar de uma só vez",
                                        enable_events=True)],
                          [sg.Checkbox("Confirmar após cada venda",
                                       key="-CONFIRMAR-", default=True, enable_events=True)]
                      ], expand_x=True
                      )],
            [sg.Frame("Tempos",
                      [
                          [sg.Text("Tempo de espera inicial (s):", size=(28, 1)),
                           sg.InputText(key="-TEMPO_INICIAL-", size=(12, 1), default_text="5",
                                        tooltip="Tempo de espera antes de iniciar",
                                        enable_events=True)],
                          [sg.Text("Tempo entre ações (s):", size=(28, 1)),
                           sg.InputText(key="-TEMPO-", size=(12, 1), default_text="0.5",
                                        tooltip="Tempo de espera entre cada ação",
                                        enable_events=True)]
                      ], expand_x=True
                      )]
        ]

        col_direita = [
            [sg.Frame("Lista de Exclusão",
                      [
                          [sg.Text("Código para excluir:", size=(18, 1)),
                           sg.InputText(key="-CODIGO_EXCLUSAO-", size=(15, 1)),
                           sg.Button(
                               "Adicionar", key="-ADICIONAR_EXCLUSAO-", size=(9, 1)),
                           sg.Button("Remover", key="-REMOVER_EXCLUSAO-", size=(9, 1))],
                          [sg.Text("Códigos excluídos:", size=(18, 1))],
                          [sg.Listbox(values=[], key="-LISTA_EXCLUSAO-", size=(None, 8),
                                      expand_x=True, expand_y=True, enable_events=True)]
                      ], expand_x=True, expand_y=True
                      )],
            [sg.Frame("Calibração de Mouse",
                      [
                          [sg.Checkbox(
                              "Usar clique do PDV", key="-USAR_CLIQUE_PDV-", default=True, enable_events=True)],
                          [sg.Text("Clique do código (PDV):", size=(22, 1)),
                           sg.Button(
                               "Calibrar", key="-CALIBRAR_CODIGO-", size=(10, 1)),
                           sg.Text("", key="-POS_CODIGO-", size=(18, 1))],
                          [sg.Text("Clique do dinheiro:", size=(22, 1)),
                           sg.Button(
                               "Calibrar", key="-CALIBRAR_DINHEIRO-", size=(10, 1)),
                           sg.Text("", key="-POS_DINHEIRO-", size=(18, 1))],
                          [sg.Text("Clique para finalizar:", size=(22, 1)),
                           sg.Button(
                               "Calibrar", key="-CALIBRAR_FINALIZAR-", size=(10, 1)),
                           sg.Text("", key="-POS_FINALIZAR-", size=(18, 1))],
                          [sg.Text("Clique para fechar:", size=(22, 1)),
                           sg.Button(
                               "Calibrar", key="-CALIBRAR_FECHAR-", size=(10, 1)),
                           sg.Text("", key="-POS_FECHAR-", size=(18, 1))]
                      ], expand_x=True
                      )]
        ]

        layout = [
            [sg.Text("Automação Softcom", font=("Arial", 16, "bold"),
                     justification="center", expand_x=True)],
            [sg.HorizontalSeparator()],
            [sg.Column(col_esquerda, vertical_alignment='top', expand_x=True, expand_y=True),
             sg.VSeparator(),
             sg.Column(col_direita, vertical_alignment='top', expand_x=True, expand_y=True)],
            [sg.HorizontalSeparator()],
            [sg.Button("Iniciar Automação", key="-INICIAR-", size=(30, 2),
                       button_color=("White", "#027F9E"), font=("Arial", 12, "bold"), expand_x=True)],
            [sg.HorizontalSeparator()],
            [sg.Output(size=(None, 12), font=("Courier", 9),
                       key="-OUTPUT-", expand_x=True, expand_y=True)],
            [sg.Text("Criado por: Jean Carlos Rodrigues Sousa | Acauã - PI",
                     justification="center", expand_x=True, font=("Arial", 8))]
        ]

        return sg.Window("Automação Softcom", layout, finalize=True, size=(1000, 750), resizable=True, keep_on_top=True)


if __name__ == "__main__":
    window = WindowAuto()
    janela = window.criar_janela()
    window.atualizar_lista_usuarios(janela)

    while True:
        event, values = janela.read()

        if event == sg.WIN_CLOSED:
            break

        elif event == '-NOVO_USUARIO-':
            novo_usuario = sg.popup_get_text("Digite o nome do novo usuário/comércio:",
                                             title="Novo Usuário", keep_on_top=True)
            if novo_usuario and novo_usuario.strip():
                novo_usuario = novo_usuario.strip()
                if ConfiguracoesUsuario.criar_usuario(novo_usuario):
                    window.atualizar_lista_usuarios(janela)
                    janela['-USUARIO-'].update(value=novo_usuario)
                    window.carregar_config_usuario(janela, novo_usuario)
                    sg.popup(f"Usuário '{novo_usuario}' criado com sucesso!",
                             title="Sucesso", keep_on_top=True)
                else:
                    sg.popup(f"Usuário '{novo_usuario}' já existe!",
                             title="Aviso", keep_on_top=True)
                    janela['-USUARIO-'].update(value=novo_usuario)
                    window.carregar_config_usuario(janela, novo_usuario)

        elif event == '-USUARIO-':
            usuario = values['-USUARIO-'].strip() if values['-USUARIO-'] else ""
            if usuario:
                window.carregar_config_usuario(janela, usuario)
            else:
                janela['-LISTA_EXCLUSAO-'].update(values=[])

        elif event == '-ADICIONAR_EXCLUSAO-':
            usuario = (values['-USUARIO-'] or "").strip()
            codigo = values['-CODIGO_EXCLUSAO-'].strip()
            if not usuario:
                sg.popup("Por favor, selecione ou crie um usuário/comércio primeiro.",
                         title="Erro", keep_on_top=True)
            elif not codigo:
                sg.popup("Por favor, informe o código para excluir.",
                         title="Erro", keep_on_top=True)
            else:
                if ConfiguracoesUsuario.adicionar_codigo_exclusao(usuario, codigo):
                    janela['-CODIGO_EXCLUSAO-'].update("")
                    window.atualizar_lista_exclusao(janela, usuario)
                    window.salvar_config_usuario(janela, usuario, values)
                else:
                    sg.popup(f"Código {codigo} já está na lista de exclusão.",
                             title="Aviso", keep_on_top=True)

        elif event == '-REMOVER_EXCLUSAO-':
            usuario = (values['-USUARIO-'] or "").strip()
            codigo_selecionado = values['-LISTA_EXCLUSAO-']
            if not usuario:
                sg.popup("Por favor, selecione ou crie um usuário/comércio primeiro.",
                         title="Erro", keep_on_top=True)
            elif not codigo_selecionado:
                codigo = values['-CODIGO_EXCLUSAO-'].strip()
                if codigo:
                    if ConfiguracoesUsuario.remover_codigo_exclusao(usuario, codigo):
                        sg.popup(f"Código {codigo} removido da lista de exclusão.",
                                 title="Sucesso", keep_on_top=True)
                        janela['-CODIGO_EXCLUSAO-'].update("")
                        window.atualizar_lista_exclusao(janela, usuario)
                        window.salvar_config_usuario(janela, usuario, values)
                    else:
                        sg.popup(f"Código {codigo} não encontrado na lista.",
                                 title="Aviso", keep_on_top=True)
                else:
                    sg.popup("Selecione um código da lista ou informe o código.",
                             title="Erro", keep_on_top=True)
            else:
                codigo = codigo_selecionado[0]
                if ConfiguracoesUsuario.remover_codigo_exclusao(usuario, codigo):
                    sg.popup(f"Código {codigo} removido da lista de exclusão.",
                             title="Sucesso", keep_on_top=True)
                    window.atualizar_lista_exclusao(janela, usuario)
                    window.salvar_config_usuario(janela, usuario, values)

        elif event == '-LISTA_EXCLUSAO-':
            if values['-LISTA_EXCLUSAO-']:
                janela['-CODIGO_EXCLUSAO-'].update(
                    values['-LISTA_EXCLUSAO-'][0])

        elif event == '-CALIBRAR_CODIGO-':
            usuario = (values.get('-USUARIO') or "").strip()
            sg.popup("Mova o mouse para a posição do código em 5 segundos...",
                     title="Calibração", keep_on_top=True)
            time.sleep(5)
            window.x_codigo, window.y_codigo = p.position()
            janela['-POS_CODIGO-'].update(
                f"X: {window.x_codigo}, Y: {window.y_codigo}")
            print(
                f"Posição do código calibrada: ({window.x_codigo}, {window.y_codigo})")
            if usuario:
                window.salvar_config_usuario(janela, usuario, values)

        elif event == '-CALIBRAR_DINHEIRO-':
            usuario = (values.get('-USUARIO') or "").strip()
            sg.popup("Mova o mouse para a posição do dinheiro em 5 segundos...",
                     title="Calibração", keep_on_top=True)
            time.sleep(5)
            window.x_dinheiro, window.y_dinheiro = p.position()
            janela['-POS_DINHEIRO-'].update(
                f"X: {window.x_dinheiro}, Y: {window.y_dinheiro}")
            print(
                f"Posição do dinheiro calibrada: ({window.x_dinheiro}, {window.y_dinheiro})")
            if usuario:
                window.salvar_config_usuario(janela, usuario, values)

        elif event == '-CALIBRAR_FINALIZAR-':
            usuario = (values.get('-USUARIO') or "").strip()
            sg.popup("Mova o mouse para a posição de finalizar em 5 segundos...",
                     title="Calibração", keep_on_top=True)
            time.sleep(5)
            window.x_finalizar, window.y_finalizar = p.position()
            janela['-POS_FINALIZAR-'].update(
                f"X: {window.x_finalizar}, Y: {window.y_finalizar}")
            print(
                f"Posição de finalizar calibrada: ({window.x_finalizar}, {window.y_finalizar})")
            if usuario:
                window.salvar_config_usuario(janela, usuario, values)

        elif event == '-CALIBRAR_FECHAR-':
            usuario = (values.get('-USUARIO') or "").strip()
            sg.popup("Mova o mouse para a posição de fechar em 5 segundos...",
                     title="Calibração", keep_on_top=True)
            time.sleep(5)
            window.x_fechar, window.y_fechar = p.position()
            janela['-POS_FECHAR-'].update(
                f"X: {window.x_fechar}, Y: {window.y_fechar}")
            print(
                f"Posição de fechar calibrada: ({window.x_fechar}, {window.y_fechar})")
            if usuario:
                window.salvar_config_usuario(janela, usuario, values)

        elif event == '-USAR_CLIQUE_PDV-':
            usuario = (values.get('-USUARIO-') or "").strip()
            usar_clique = values['-USAR_CLIQUE_PDV-']
            janela['-CALIBRAR_CODIGO-'].update(disabled=not usar_clique)
            if not usar_clique:
                janela['-POS_CODIGO-'].update("")
            if usuario:
                window.salvar_config_usuario(janela, usuario, values)

        elif event in ['-VALOR_MAX-', '-MARGEM-', '-PORC_ESTOQUE-', '-ESTOQUE_MIN-',
                       '-TEMPO_INICIAL-', '-TEMPO-', '-CONFIRMAR-', '-QTD_MAX_ITEM-', '-QTD_MAX_POR_VEZ-']:
            usuario = (values.get('-USUARIO-') or "").strip()
            if usuario:
                window.salvar_config_usuario(janela, usuario, values)

        elif event == '-INICIAR-':
            if not values['-PATH-']:
                sg.popup("Por favor, selecione o arquivo Excel.",
                         title="Erro", keep_on_top=True)
                continue

            usar_clique_pdv = values.get('-USAR_CLIQUE_PDV-', True)

            if usar_clique_pdv and (window.x_codigo is None or window.y_codigo is None):
                sg.popup("Por favor, calibre a posição do código primeiro.",
                         title="Erro", keep_on_top=True)
                continue

            if window.x_dinheiro is None or window.y_dinheiro is None:
                sg.popup("Por favor, calibre a posição do dinheiro primeiro.",
                         title="Erro", keep_on_top=True)
                continue

            if window.x_finalizar is None or window.y_finalizar is None:
                sg.popup("Por favor, calibre a posição de finalizar primeiro.",
                         title="Erro", keep_on_top=True)
                continue

            if window.x_fechar is None or window.y_fechar is None:
                sg.popup("Por favor, calibre a posição de fechar primeiro.",
                         title="Erro", keep_on_top=True)
                continue

            usuario = (values.get('-USUARIO-') or '').strip()

            if usuario:
                window.salvar_config_usuario(janela, usuario, values)

            try:
                auto = AutoSoftcom(
                    path=values['-PATH-'],
                    tempo_espera_inicial=values.get('-TEMPO_INICIAL-', '5'),
                    tempo_espera=values['-TEMPO-'],
                    valor_max_venda=values['-VALOR_MAX-'],
                    x_codigo=window.x_codigo if usar_clique_pdv else 0,
                    y_codigo=window.y_codigo if usar_clique_pdv else 0,
                    x_dinheiro=window.x_dinheiro,
                    y_dinheiro=window.y_dinheiro,
                    x_finalizar=window.x_finalizar,
                    y_finalizar=window.y_finalizar,
                    x_fechar=window.x_fechar,
                    y_fechar=window.y_fechar,
                    porcentagem_estoque=values['-PORC_ESTOQUE-'],
                    estoque_minimo=values['-ESTOQUE_MIN-'],
                    margem_venda=values['-MARGEM-'],
                    confirmar_venda=values['-CONFIRMAR-'],
                    usar_clique_pdv=usar_clique_pdv,
                    usuario=usuario if usuario else None,
                    quantidade_max_item=values.get('-QTD_MAX_ITEM-', '200'),
                    quantidade_max_por_vez=values.get(
                        '-QTD_MAX_POR_VEZ-', '99')
                )
                auto.executar()
            except Exception as e:
                sg.popup(f"Erro: {e}", title="Erro", keep_on_top=True)
                print(f"Erro detalhado: {e}")

    janela.close()
