# Padrão do topo (header) dos apps Streamlit da TAG

Referência: topo do Mesa de Boletas (`tema.py`, função `topo()`). Use esta especificação em todos os projetos, para que o cabeçalho fique idêntico.

## 1. Estrutura

Uma faixa horizontal com:
- **À esquerda:** título e subtítulo (dentro de uma `div`).
- **À direita:** logo TAG branca (`logo_tag.png`).

Não usar selo de texto "TAG Investimentos" quando a logo estiver presente, e não colocar logo na esquerda. Se o arquivo da logo faltar, o topo pode mostrar o selo de texto como alternativa.

## 2. Especificação visual

| Item | Valor |
|---|---|
| Fundo | Gradiente `120deg`: `#4A0A1E` (0%) → `#6A0F2B` (60%) → `#8A1B3D` (100%) |
| Texto | Branco |
| Borda inferior | 3 px, dourado `#B8924A` |
| Cantos | Raio `16px` |
| Sombra | `0 6px 20px rgba(74, 10, 30, .28)` |
| Espaçamento interno | `1rem 1.6rem` |
| Espaço entre blocos | `gap: 1.4rem` |
| Margem inferior | `1.2rem` |
| Fonte | Inter (pesos 400 a 700), com `system-ui` de reserva |
| Título (`.marca`) | `1.8rem`, peso 700, entrelinha `1.1`, `letter-spacing: -.01em`, margem 0 |
| Subtítulo (`.sub`) | `.92rem`, entrelinha `1.3`, cor `#E9CFD7`, margem `.2rem 0 0 0` |
| Logo | Altura `52px`, largura automática, `margin-left: auto` (empurra para a direita), sem encolher |
| Largura do conteúdo da página | `.block-container` com `max-width: 1320px` |

A logo é o `logo_tag.png`, branco com fundo transparente, recortado sem borda vazia (aprox. 450×161 px). Entra no HTML em base64.

## 3. CSS

```css
.block-container { padding-top: 1rem; padding-bottom: 3rem; max-width: 1320px; }

.topo {
    display: flex; align-items: center; gap: 1.4rem;
    background: linear-gradient(120deg, #4A0A1E 0%, #6A0F2B 60%, #8A1B3D 100%);
    color: #fff; padding: 1rem 1.6rem; border-radius: 16px; margin-bottom: 1.2rem;
    box-shadow: 0 6px 20px rgba(74, 10, 30, .28);
    border-bottom: 3px solid #B8924A;
}
.topo .marca { font-family: 'Inter', system-ui, sans-serif; font-size: 1.8rem; font-weight: 700;
               letter-spacing: -.01em; margin: 0; color: #fff; line-height: 1.1; }
.topo .sub   { margin: .2rem 0 0 0; font-size: .92rem; line-height: 1.3; color: #E9CFD7; }
.topo img    { height: 52px; width: auto; margin-left: auto; flex-shrink: 0; }
```

Fonte Inter, no início do CSS:

```css
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
```

## 4. HTML

```html
<div class="topo">
  <div>
    <p class="marca">TÍTULO DO APP</p>
    <p class="sub">Subtítulo curto, de uma linha</p>
  </div>
  <img src="data:image/png;base64,..." alt="TAG Investimentos">
</div>
```

## 5. Função Python (Streamlit)

```python
import base64
from pathlib import Path
import streamlit as st

LOGO = Path(__file__).resolve().with_name("logo_tag.png")

def topo(titulo: str, subtitulo: str) -> None:
    try:
        b64 = base64.b64encode(LOGO.read_bytes()).decode()
    except OSError:
        b64 = ""
    img = f'<img src="data:image/png;base64,{b64}" alt="TAG Investimentos">' if b64 else ""
    st.markdown(
        f'<div class="topo"><div><p class="marca">{titulo}</p><p class="sub">{subtitulo}</p></div>'
        f'{img or "<span class=selo>TAG Investimentos</span>"}</div>',
        unsafe_allow_html=True,
    )
```

Chamar logo depois do `st.set_page_config` e do CSS da página:

```python
topo("Mesa de Boletas", "Conferência, liquidação e controle de erros em um só lugar")
```

## 6. Diferenças encontradas entre os dois projetos e o que ajustar

| Ponto | Extratos mensais | Mesa de Boletas | Ajuste no padrão |
|---|---|---|---|
| Altura | Mais alto | Compacto | Padding `1rem 1.6rem` nos dois |
| Subtítulo | 2 linhas, entrelinha grande | 1 linha | `.92rem`, entrelinha `1.3` e texto curto |
| Título | Maior | Menor | `1.8rem`, peso 700 |
| Logo | Pequena, colada na borda direita | Maior, com respiro | Altura `52px` e padding direito de `1.6rem` |
| Largura | Vai até a borda da tela | Para na largura do conteúdo | `max-width: 1320px` no `.block-container` |

## 7. Recomendações para o texto

- **Título:** curto, até uns 30 caracteres (ex.: "Mesa de Boletas", "Solicitação de extratos").
- **Subtítulo:** uma linha, uns 60 caracteres no máximo (estimativa minha; depende da largura da tela). Exemplo: "Gere os e-mails de extratos direto no Outlook e envie".
- Se o subtítulo for maior que isso, ele quebra em duas linhas. O padrão continua válido, mas o topo fica mais alto.

## 8. Checklist para ajustar outro projeto

- [ ] Fonte Inter carregada
- [ ] CSS de `.block-container` com `max-width: 1320px`
- [ ] CSS `.topo`, `.marca`, `.sub` e `img` copiados
- [ ] `logo_tag.png` recortado na pasta do projeto
- [ ] Função `topo()` no lugar do HTML antigo
- [ ] Subtítulo de uma linha
- [ ] Reiniciar o app (o Streamlit não recarrega o `tema.py` sozinho)
