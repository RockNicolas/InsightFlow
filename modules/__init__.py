import pandas as pd
import matplotlib.pyplot as plt
import glob
import os

def generate_professional_dashboard(folder):
    # 1. Busca e junta os dados
    files = glob.glob(os.path.join(folder, "history_*.csv"))
    if len(files) < 2: return None

    df_total = pd.concat([pd.read_csv(f) for f in files])
    
    # --- O SEGREDO: Criar um eixo X numérico para a linha inclinar ---
    ordem_semanas = sorted(df_total['Semana'].unique())
    mapa_x = {semana: i for i, semana in enumerate(ordem_semanas)}
    df_total['x_coords'] = df_total['Semana'].map(mapa_x)
    df_total = df_total.sort_values(['machine', 'x_coords'])

    # 2. Estilo Dark High-Tech
    plt.style.use('dark_background')
    fig = plt.figure(figsize=(20, 14), facecolor='#0B0F19')
    grid = fig.add_gridspec(2, 2, height_ratios=[1, 0.3], hspace=0.3, wspace=0.2)

    # Título Neon
    fig.text(0.12, 0.95, "INSIGHTFLOW | ANALYTICS AREA CHART", fontsize=28, fontweight='bold', color='#00F2FF')

    # Separação Máquinas vs Veículos
    df_maq = df_total[df_total['machine'].str.contains('MC|LOC', case=False, na=False)]
    df_vei = df_total[~df_total['machine'].str.contains('MC|LOC', case=False, na=False)]

    secoes = [(df_maq, "MÁQUINAS (HORAS)", 0), (df_vei, "VEÍCULOS (KM)", 1)]

    for data, title, pos in secoes:
        ax = fig.add_subplot(grid[0, pos])
        ax.set_facecolor('#0B0F19')
        
        # Filtra as 5 principais para o visual ficar limpo como na foto
        top_5 = data.groupby('machine')['hours'].sum().nlargest(5).index
        
        for name in top_5:
            subset = data[data['machine'] == name]
            
            # --- DESENHA A LINHA (Sem Bolas, Linha Grossa) ---
            line, = ax.plot(subset['x_coords'], subset['hours'], label=name, linewidth=6, alpha=0.9)
            
            # --- PREENCHIMENTO DE ÁREA (O efeito que você quer) ---
            ax.fill_between(subset['x_coords'].astype(float), subset['hours'].astype(float), 
                            color=line.get_color(), alpha=0.2)

        # Ajusta o Eixo X para mostrar os nomes das semanas com espaço entre elas
        ax.set_xticks(range(len(ordem_semanas)))
        ax.set_xticklabels(ordem_semanas)
        
        ax.grid(True, linestyle='--', alpha=0.1, color='#64748B')
        ax.set_title(title, fontsize=18, fontweight='bold', color='white', pad=20)
        ax.legend(loc='upper left', bbox_to_anchor=(1, 1), frameon=False, fontsize=10)

    # --- RODAPÉ COM TEXTO COMPARATIVO ---
    ax_txt = fig.add_subplot(grid[1, :])
    ax_txt.axis('off')
    
    s1, s2 = ordem_semanas[0], ordem_semanas[1]
    h1 = df_maq[df_maq['Semana'] == s1]['hours'].sum()
    h2 = df_maq[df_maq['Semana'] == s2]['hours'].sum()
    
    status = "MAIOR" if h2 > h1 else "MENOR"
    texto = f"NA SEMANA {s2} A PRODUÇÃO TOTAL FOI {status} QUE NA SEMANA {s1}."
    
    ax_txt.text(0.5, 0.5, texto, ha='center', va='center', fontsize=18, family='monospace',
                color='#00F2FF', fontweight='bold', bbox=dict(facecolor='#161B22', edgecolor='#00F2FF', pad=20))

    plt.subplots_adjust(left=0.1, right=0.88, top=0.9, bottom=0.1)
    
    output_path = os.path.join(folder, "DASHBOARD_AREA_FINAL.png")
    plt.savefig(output_path, dpi=300, facecolor='#0B0F19')
    plt.close()
    return "DASHBOARD_AREA_FINAL.png"