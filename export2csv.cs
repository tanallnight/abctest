using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using InspectionWare;

namespace Export2CSV {
	
	//Comment

    public class Executable : UserDefinedExecutable {

        int _nSubset = 0;
        public int SubsetIndex { set { _nSubset = value; } }

        bool _bConvertData = true;
        public bool ExportConvertedData { set { _bConvertData = value; } }

        string _sFilePath = "";
        public string SaveFolder { set { _sFilePath = value; } }

        string _sFileName = "";
        public string FileName { set { _sFileName = value; } }

        bool _bLaunchOnExport = false;
        public bool LaunchOnExport { set { _bLaunchOnExport = value; } }

        string _sSubsetName = "";

        IWNode dataSet;
        IWDataObject subset;
        IWDataTable outputTable;

        private void Initialize() {
            dataSet = (IWNode)_Node.GetSourceObject(0);
            subset = dataSet.GetDataObject(_nSubset);

            _sSubsetName = subset.GetStringProp("Name");
        }

        public override bool Execute()  {
            Initialize();

            int _nAxes = subset.AxisCount;
            List<DataAxis> _lAxes = new List<DataAxis>();
            for (int i = 0; i < _nAxes; i++) {
                _lAxes.Add(new DataAxis(subset.GetAxisType(i), subset.GetAxisUnits(i),
                    (float) subset.GetAxisRes(i), (float) subset.GetAxisStart(i), subset.GetAxisPoints(i)));
            }

            string _sFullPath = "";
            if (string.IsNullOrEmpty(_sFilePath) || string.IsNullOrEmpty(_sFileName)) {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                saveFileDialog.Filter = "CSV File (*.csv)|*.csv";
                saveFileDialog.ShowDialog();
                _sFullPath = saveFileDialog.FileName;
            } else {
                if (!_sFilePath.EndsWith("/"))
                    _sFilePath = _sFilePath + "/";
                _sFullPath = _sFilePath + _sFileName + ".csv";
            }

            string _sHeader = GetHeader(_lAxes);
            string _sCSVConverted = "";
            
            using (StreamWriter streamWriter = new StreamWriter(_sFullPath)) {
                streamWriter.WriteLine(_sHeader);
                if (_lAxes.Count == 1) {
                    for (int i = 0; i < _lAxes[0]._nAxisPoints; i++) {
                        float _dAxisPoint = _lAxes[0]._dAxisStart + (i * _lAxes[0]._dAxisRes);
                        float _dDataPoint = (float) subset.GetDataPoint(i, 0, 0, 0, _bConvertData);

                        _sCSVConverted = _dAxisPoint + "," + _dDataPoint;

                        streamWriter.WriteLine(_sCSVConverted);
                    }
                } else if (_lAxes.Count == 2) {
                    for (int i = 0; i < _lAxes[1]._nAxisPoints; i++) {
                        for (int j = 0; j < _lAxes[0]._nAxisPoints; j++) {
                            float _dIndex = _lAxes[1]._dAxisStart + (i * _lAxes[1]._dAxisRes);
                            float _dScan = _lAxes[0]._dAxisStart + (j * _lAxes[0]._dAxisRes);
                            float _dDataPoint = (float) subset.GetDataPoint(j, i, 0, 0, _bConvertData);

                            _sCSVConverted = _dIndex + "," + _dScan + "," + _dDataPoint;

                            streamWriter.WriteLine(_sCSVConverted);
                        }
                    }
                } else if (_lAxes.Count == 3) {
                    for (int i = 0; i < _lAxes[2]._nAxisPoints; i++) {
                        for (int j = 0; j < _lAxes[1]._nAxisPoints; j++) {
                            float _dIndex = _lAxes[2]._dAxisStart + (i * _lAxes[2]._dAxisRes);
                            float _dScan = _lAxes[1]._dAxisStart + (j * _lAxes[1]._dAxisRes);
                            float[] _fAScan = new float[_lAxes[0]._nAxisPoints];
                            subset.GetData(ref _fAScan, i, j);

                            _sCSVConverted = _dIndex + "," + _dScan;
                            streamWriter.Write(_sCSVConverted);
                            for (int k = 0; k < _fAScan.Length; k++) {
                                streamWriter.Write("," + _fAScan[k]);
                            }

                            streamWriter.WriteLine();
                        }
                    }
                }
            }

            if (_bLaunchOnExport) {
                Process.Start(_sFullPath);
            }

            return true;
        }

        private string GetHeader(List<DataAxis> _lAxes) {
            string _sHeader = "";
            if (_lAxes.Count == 1) {
                _sHeader = _lAxes[0]._sAxisType + "," + "Data";
            } else if (_lAxes.Count == 2) {
                _sHeader = _lAxes[1]._sAxisType + "," + _lAxes[0]._sAxisType + "," + "Data";
            } else if (_lAxes.Count == 3) {
                _sHeader = _lAxes[2]._sAxisType + "," + _lAxes[1]._sAxisType;
                for (int i = 0; i < _lAxes[0]._nAxisPoints; i++) {
                    float _dDataPoint = _lAxes[0]._dAxisStart + (i * _lAxes[0]._dAxisRes);
                    _sHeader = _sHeader + "," + _dDataPoint;
                }
            }
            return _sHeader;
        }

        private class DataAxis {
            public string _sAxisType;
            public string _sAxisUnits;
            public float _dAxisRes;
            public float _dAxisStart;
            public int _nAxisPoints;

            public DataAxis(string _sAxisType, string _sAxisUnits, float _dAxisRes, float _dAxisStart, int _nAxisPoints) {
                this._sAxisType = _sAxisType;
                this._sAxisUnits = _sAxisUnits;
                this._dAxisRes = _dAxisRes;
                this._dAxisStart = _dAxisStart;
                this._nAxisPoints = _nAxisPoints;
            }

            public override string ToString() {
                return "Axis Type: " + _sAxisType + "\n" +
                    "Axis Units: " + _sAxisUnits + "\n" +
                    "Axis Start: " + _dAxisStart + "\n" +
                    "Axis Resolution: " + _dAxisRes + "\n" +
                    "Number Of Points: " + _nAxisPoints;
            }
        }

    }
}
