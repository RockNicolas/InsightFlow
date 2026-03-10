import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import glob
import os

def generate_monthly_dashboard(output_folder):
    """Lê todos os CSVs da pasta outputs e gera um gráfico comparativo."""
    