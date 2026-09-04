import customtkinter as ctk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


def create_pie_chart(tab_name, sum_1, sum_2, sum_3, sum_4):
    """Clears the statistics tab and builds a pie chart of earnings distribution."""
    # Close all previous figures to prevent memory leaks
    plt.close('all')
    
    # Clear the statistics tab from old charts or text labels
    for widget in tab_name.winfo_children():
        widget.destroy()

    labels = []     # Chart sector labels
    sizes = []      # Value sizes (earnings per report)

    # Collect reports where earnings are strictly greater than zero
    if sum_1 > 0:
        labels.append("Работа с РТ")
        sizes.append(sum_1)
    if sum_2 > 0:
        labels.append("Приемка")
        sizes.append(sum_2)
    if sum_3 > 0:
        labels.append("Пики")
        sizes.append(sum_3)
    if sum_4 > 0:
        labels.append("Зачистки")
        sizes.append(sum_4)

    # Display a text notification if there is no data to plot
    if not sizes:
        no_data_label = ctk.CTkLabel(tab_name, text="Нет данных для диаграммы", font=("Arial", 16))
        no_data_label.pack(pady=50)
        return None

    # Apply dark background style to matplotlib
    plt.style.use("dark_background")
    fig, ax = plt.subplots(figsize=(5, 4), dpi=100)
    
    # Color palette for pie chart sectors
    colors = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728"]

    # Render the pie chart
    ax.pie(
        sizes, 
        labels=labels, 
        autopct="%1.1f%%", 
        startangle=140, 
        colors=colors[:len(sizes)], 
        radius=0.9, 
        textprops={"fontsize": 11}
    )
    ax.axis("equal")  # Ensure the pie chart is rendered as a circle

    # Match chart background color with CustomTkinter dark theme (#2b2b2b)
    fig.patch.set_facecolor("#2b2b2b")

    # Embed the matplotlib chart into the CustomTkinter frame
    canvas = FigureCanvasTkAgg(fig, master=tab_name)
    canvas.draw()
    
    widget = canvas.get_tk_widget()
    widget.pack(pady=10, padx=10, fill="both", expand=True)        
    
    return canvas
