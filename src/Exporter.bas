Attribute VB_Name = "Exporter"
' # ------------------------------------------------------------------------------
' # Name:        Exporter.bas
' # Purpose:     Helper module for launching the CSV Exporter add-in
' #
' # Original Author:  Brian Skinn
' #                     bskinn@alum.mit.edu
' #
' # Modifications:    Shep Bivens
' #   - Added tab delimiter support
' #   - Added Excel-compatible quoting modes
' #   - Improved number formatting behavior
' #   - Added choice of output encoding
' #   - UI and usability enhancements
' #
' # Created:     24 Jan 2016
' # Modified:    01 Feb 2026
' #
' # Copyright:
' #    (c) Brian Skinn 2016-2022
' #    Modifications (c) Shep Bivens 2026
' #
' # License:   The MIT License; see "LICENSE.txt" for full license terms.
' #
' # Original:  http://www.github.com/bskinn/excel-csvexporter
' # Fork:      https://github.com/Shep-Bivens/excel-csvexporter
' #
' # ------------------------------------------------------------------------------

Option Explicit

Public Sub showCSVExporterForm()
    UFExporter.RefreshForShow
    UFExporter.Show
End Sub

Public Sub Ribbon_ShowExporter(control As IRibbonControl)
    showCSVExporterForm
End Sub

