
# 📊 Matplotlib Tutorial

Matplotlib is the foundational data visualization library in Python. It provides a flexible framework for creating static, animated, and interactive plots in a variety of formats.

## 📐 Plotting with Pyplot API

The `pyplot` interface mimics MATLAB and is best for quick and simple plotting.

- 🔹 `plot()` draws line charts
- 🔹 `scatter()` for individual points
- 🔹 `bar()` and `barh()` for bar charts
- 🔹 `hist()` for histograms
- 🔹 `imshow()` for image data

```python
import matplotlib.pyplot as plt
x = [1, 2, 3, 4, 5]
y = [1, 4, 9, 16, 25]
plt.figure()
plt.plot(x, y, marker='o', linestyle='--', color='r')
plt.title("Line Chart Example")
plt.xlabel("X Axis")
plt.ylabel("Y Axis")
plt.show()
```

## 🧱 Object-Oriented API

Use the object-oriented interface for fine-grained control.

- 🔹 Explicit `Figure` and `Axes` objects
- 🔹 Ideal for multi-plot layouts
- 🔹 Better for reusable, maintainable code

```python
fig, ax = plt.subplots()
x = np.linspace(0, 2 * np.pi, 100)
ax.plot(x, np.sin(x))
ax.set_title("Sine Function")
ax.set_xlabel("x")
ax.set_ylabel("sin(x)")
plt.show()
```

## 🎨 Customization

Matplotlib allows nearly every aspect of the plot to be customized.

- 🔹 Change colors, markers, linestyles
- 🔹 Annotate text or labels
- 🔹 Add grids, legends
- 🔹 Modify layout with `tight_layout()`

```python
plt.plot(x, y, color='purple', linewidth=2, linestyle='-.')
plt.title("Customized Plot")
plt.grid(True)
plt.legend(["sin(x)"])
plt.annotate("Peak", xy=(1.57, 1), xytext=(3, 1.2),
             arrowprops=dict(facecolor='black'))
plt.tight_layout()
plt.show()
```

## 🧮 Working with Subplots

Use `subplots()` to create layouts with multiple plots.

```python
fig, axs = plt.subplots(2, 2, figsize=(8, 6))
x = np.linspace(0, 10, 100)

axs[0, 0].plot(x, np.sin(x))
axs[0, 1].plot(x, np.cos(x))
axs[1, 0].plot(x, np.tan(x))
axs[1, 1].plot(x, -np.sin(x))

plt.tight_layout()
plt.show()
```

## 🌈 Colormaps and Heatmaps

Use colormaps for scalar field or image-like visualizations.

```python
data = np.random.rand(10, 10)
plt.imshow(data, cmap='viridis')
plt.colorbar(label='Value')
plt.title("Heatmap")
plt.show()
```

## 💾 Saving Figures

Save your figures using `savefig()`.

```python
plt.plot(x, y)
plt.title("Exported Plot")
plt.savefig("sine_plot.png", dpi=300, transparent=True)
plt.close()
```

## 📚 Summary

Matplotlib is a powerful and customizable plotting library.

- 🔹 Use `pyplot` for quick charts
- 🔹 Use OO API for complex layouts
- 🔹 Customize every element of the plot
- 🔹 Integrates well with Pandas, NumPy
- 🔹 Ideal for both EDA and publications
