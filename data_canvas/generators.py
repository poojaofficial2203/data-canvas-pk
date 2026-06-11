import matplotlib.pyplot as plt
import io
from PIL import Image

class ChartGenerator:

    def histogram(self, data, title="Histogram"):

        plt.figure(figsize=(6,4))
        plt.hist(data)

        plt.title(title)

        buffer = io.BytesIO()

        plt.savefig(buffer)
        plt.close()

        buffer.seek(0)

        return Image.open(buffer)

    def bar_chart(self, data, title="Bar Chart"):

        plt.figure(figsize=(6,4))

        data.value_counts().head(10).plot(kind="bar")

        plt.title(title)

        buffer = io.BytesIO()

        plt.savefig(buffer)
        plt.close()

        buffer.seek(0)

        return Image.open(buffer)

    def scatter_plot(self, df, x_col, y_col, title="Scatter Plot"):

        plt.figure(figsize=(6,4))

        plt.scatter(df[x_col], df[y_col])

        plt.title(title)

        buffer = io.BytesIO()

        plt.savefig(buffer)
        plt.close()

        buffer.seek(0)

        return Image.open(buffer)
