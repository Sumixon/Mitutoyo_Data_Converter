# Mitutoyo Data Converter

A modern desktop application for converting measurement data from the Mitutoyo SJ-412 device from .txt format to Excel for Windows.

## 📋 Features

- ✅ **Import TXT files** from the Mitutoyo SJ-412 measuring device
- ✅ **Automatic processing** of measurement data
- ✅ **Export to Excel** format (.xlsx)
- ✅ **Support for all roughness parameters** (Ra, Rz, Rq, Rp, Rv, etc.)
- ✅ **Modern GUI** with elegant design
- ✅ **Batch processing** - handle multiple files at once
- ✅ **Intuitive user interface**

## 🖥️ System Requirements

- **Operating System:** Windows 10/11
- **Python:** 3.8 or newer
- **RAM:** At least 4GB
- **Disk Space:** 100MB for the app + space for data

## 🚀 Installation

### Option 1: Run from source

1. **Clone the repository:**
   ```bash
   git clone https://github.com/Sumixon/mitutoyo-converter.git
   cd mitutoyo-converter
   ```

2. **Create a virtual environment:**
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the application:**
   ```bash
   python main.pyw
   ```

### Option 2: Create standalone EXE

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=icon.ico main.pyw
```

The resulting EXE file will be in the `dist/` folder.

## 🚀 Quick Start

1. **Start the app** - `python main.pyw`
2. **Import files** - click "📂 Import Files"
3. **Select TXT files** from the Mitutoyo SJ-412 device
4. **Review data** in the table
5. **Export to Excel** - click "📊 Export to Excel"
6. **Save the file** to your desired location

## 📊 Supported Parameters

| Parameter | Unit | Description |
|-----------|------|-------------|
| Ra | μm | Arithmetic average roughness |
| Rz | μm | Mean roughness depth |
| Rq | μm | RMS roughness |
| Rp | μm | Maximum profile peak height |
| Rv | μm | Maximum profile valley depth |
| Rsk | μm | Profile skewness |
| Rku | μm | Profile kurtosis |
| Rc | μm | Mean height of profile elements |
| RPc | /cm | Peak count per cm |
| RSm | μm | Mean spacing of profile elements |
| RDq | μm | Root mean square slope |
| Rmr | % | Material ratio of the bearing length curve |
| Rdc | μm | Profile height |
| Rt | μm | Total height of the profile |
| Rz1max | μm | Maximum roughness height |
| Rk | μm | Core roughness depth |
| Rpk | μm | Reduced peak height |
| Rvk | μm | Reduced valley depth |
| Mr1 | % | Material ratio 1 |
| Mr2 | % | Material ratio 2 |
| A1 | - | Area above the core |
| A2 | - | Area below the core |

## 🔧 Technical Details

- **Framework:** Tkinter with modern ttk styling
- **Data processing:** Pandas for data manipulation
- **Excel export:** OpenPyXL for .xlsx file creation
- **GUI Style:** Modern flat design with Material Design elements
- **File handling:** UTF-8 encoding with error handling support
- **Architecture:** Object-oriented design with modular structure

## 📋 Input File Format

The app expects TXT files from the Mitutoyo SJ-412 in the following structure:

```
 //Header
 Date;2025-01-01;
 Time;10:30:15;

 //CalcResult  
 Ra;1.234;μm
 Rz;5.678;μm
 Rq;1.456;μm
 ...

 //Condition-A
 Cutoff;0.8;mm
 Speed;0.5;mm/s
 ...
```

## 🐛 Troubleshooting

### Common issues:

**App does not start:**
- Check that Python 3.8+ is installed
- Verify all dependencies: `pip install -r requirements.txt`

**Error reading TXT file:**
- Ensure the file is in the correct Mitutoyo SJ-412 format
- Check that file encoding is UTF-8

**Excel export not working:**
- Check write permissions for the target folder
- Make sure the target Excel file is not open

**Slow processing:**
- For a large number of files, consider processing in smaller batches
- Check available RAM

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Coding standards:
- Use Python PEP 8
- Add docstrings for all functions
- Write tests for new features

## 📄 License

Distributed under the MIT License. See `LICENSE` for more information.

## 👨‍💻 Author

**Roman Denev (Sumixon)**
- GitHub: [@Sumixon](https://github.com/Sumixon)
- Email: romna.denev@gmail.com

## 🙏 Acknowledgements

- [Python Software Foundation](https://www.python.org/) for a great programming language
- [Pandas](https://pandas.pydata.org/) for powerful data processing
- [OpenPyXL](https://openpyxl.readthedocs.io/) for Excel export functionality
- [Tkinter](https://docs.python.org/3/library/tkinter.html) for the GUI framework

## 📈 Changelog

### v2.0.0 (2025-01-01)
- ✅ Completely redesigned modern UI
- ✅ Improved TXT file parser with better error handling
- ✅ Extended support for all roughness parameters
- ✅ Optimized processing of large files
- ✅ Added tabs for better organization

### v1.0.0 (2024-12-01)
- ✅ Initial version of the application
- ✅ Basic import/export functionality
- ✅ Tkinter GUI with basic design

## 🔗 Useful links

- [Mitutoyo SJ-412 Manual](https://mitutoyo.com/)
- [Python Documentation](https://docs.python.org/3/)
- [Pandas Documentation](https://pandas.pydata.org/docs/)
- [Tkinter Tutorial](https://docs.python.org/3/library/tkinter.html)

---

**Made with ❤️ for precise surface roughness measurement**
