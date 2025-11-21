## 🧮 NumPy Tutorial: A Foundational Guide for Numerical Computing in Python

NumPy (Numerical Python) is a fundamental package for scientific computing in Python. 
It provides powerful multi-dimensional arrays, broadcasting functions, linear algebra 
capabilities, and integration with C/C++ code. This tutorial introduces key concepts 
with examples to help you get started.


### 📦 Getting Started with NumPy

NumPy is installed via `pip` or comes pre-installed with Anaconda. The library is typically imported as `np`.

- Provides ndarray, a fast, space-efficient array object.
- Enables vectorized operations for high performance.
- Interfaces seamlessly with other libraries like SciPy, Pandas, and Scikit-learn.
- Useful for numerical simulations, data analysis, and ML pipelines.


```
import numpy as np

# Creating a simple array
arr = np.array([1, 2, 3])
print(arr)
```

### 📐 Array Creation

NumPy provides functions to create arrays of different shapes and values.

- `np.array` creates arrays from Python lists or tuples.
- `np.zeros`, `np.ones`, `np.empty` create arrays with default values.
- `np.arange`, `np.linspace` for sequences.
- `np.eye` for identity matrices.


```
zeros = np.zeros((2, 3))
ones = np.ones((2, 3))
range_array = np.arange(0, 10, 2)
linspace_array = np.linspace(0, 1, 5)
identity = np.eye(3)
```

### 🔄 Array Manipulation

Reshape, flatten, concatenate, and split arrays using these techniques:

- `reshape()` changes shape without changing data.
- `flatten()` turns an N-D array into 1-D.
- `concatenate()` joins arrays along an axis.
- `split()` or `hsplit()`, `vsplit()` divide arrays.


```
x = np.array([[1, 2, 3], [4, 5, 6]])
reshaped = x.reshape((3, 2))
flattened = x.flatten()
stacked = np.concatenate((x, x), axis=0)
```

### ➕ Arithmetic and Broadcasting

NumPy supports element-wise arithmetic operations and broadcasting.

- Arithmetic (+, -, *, /, **) is element-wise.
- Broadcasting aligns arrays with different shapes.
- Mathematical functions: `np.sqrt`, `np.exp`, `np.sin`, etc.


```
a = np.array([1, 2, 3])
b = np.array([4, 5, 6])

add = a + b
square = a ** 2
sin_vals = np.sin(a)
```

### 🧠 Indexing, Slicing, Masking

Flexible access and modification of array elements.

- Similar to Python lists but extended for multi-dimensional arrays.
- Slicing syntax: `arr[start:stop:step]`
- Boolean masking to filter elements.
- Fancy indexing with arrays of indices.


```
arr = np.array([[10, 20, 30], [40, 50, 60]])
sub_array = arr[:, 1]      # Second column
mask = arr > 30
filtered = arr[mask]       # Values > 30
```

### 🧮 Aggregation & Statistics

Efficient summarization of data.

- `sum`, `mean`, `std`, `var`, `min`, `max`
- Specify `axis` for row/column-wise aggregation
- `np.argmin`, `np.argmax` for index of min/max


```
arr = np.array([[1, 2], [3, 4]])
total = arr.sum()
column_means = arr.mean(axis=0)
row_max = arr.max(axis=1)
```

### 🔁 Random Number Generation

NumPy’s `random` module is used for simulations and stochastic algorithms.

- `np.random.rand` for uniform distribution
- `np.random.randn` for standard normal distribution
- `np.random.randint` for integers
- `np.random.seed` ensures reproducibility


```
np.random.seed(0)
rand_uniform = np.random.rand(2, 2)
rand_normal = np.random.randn(3)
rand_ints = np.random.randint(0, 10, size=(2, 3))
```

### 🧾 Linear Algebra

The `linalg` module provides support for solving linear equations, decompositions, and more.

- `np.dot`, `@` for dot products
- `np.linalg.inv`, `np.linalg.det` for matrix inverse and determinant
- `np.linalg.eig` for eigenvalues/vectors


```
A = np.array([[1, 2], [3, 4]])
B = np.array([[5, 6], [7, 8]])
product = A @ B
inv_A = np.linalg.inv(A)
eigenvalues, eigenvectors = np.linalg.eig(A)
```

### 🧼 Practical Tips

Best practices to avoid bugs and optimize performance.

- Use `astype()` to change types.
- Use `np.where()` for conditionals.
- Use `.copy()` to avoid unintentional reference issues.


```
arr = np.array([1.7, 2.5, 3.3])
int_arr = arr.astype(int)
conditioned = np.where(arr > 2, 1, 0)
arr_copy = arr.copy()
```

### ✅ Summary

NumPy is the bedrock of numerical computing in Python. With arrays, broadcasting, 
and high-performance operations, it enables efficient computation workflows essential 
for data science, ML, and scientific applications.

> Tip: Combine NumPy with Pandas for labeled data handling, or SciPy for advanced numerical methods.

