import numpy as np
from PIL import Image
import matplotlib.pyplot as plt
from matplotlib import cm
from matplotlib.colors import Normalize

def extract_grid_values(img_path, n_rows, n_cols, cmap=cm.viridis, vmin=0, vmax=1):
    # Load and normalize image
    img = Image.open(img_path).convert('RGB')
    img_data = np.array(img) / 255.0  # normalize to 0–1
    height, width, _ = img_data.shape

    # Set up value-color mapping
    norm = Normalize(vmin=vmin, vmax=vmax)
    values = np.linspace(vmin, vmax, 1000)
    colormap = cmap(norm(values))[:, :3]

    # Compute grid cell size
    cell_height = height // n_rows
    cell_width = width // n_cols

    def color_to_value(color):
        diff = np.linalg.norm(colormap - color, axis=1)
        return values[np.argmin(diff)]

    # Extract values from each cell (use center pixel)
    result = np.zeros((n_rows, n_cols))
    for i in range(n_rows):
        for j in range(n_cols):
            cy = int((i + 0.5) * cell_height)
            cx = int((j + 0.5) * cell_width)
            rgb = img_data[cy, cx, :3]
            result[i, j] = color_to_value(rgb)

    return result

values = extract_grid_values("energy duty.png", n_rows=10, n_cols=8, cmap=cm.viridis, vmin=10.0, vmax=50.0)

plt.imshow(values, cmap='viridis')
plt.colorbar(label='Extracted Value')
plt.title("Extracted Grid Values")
plt.show()
