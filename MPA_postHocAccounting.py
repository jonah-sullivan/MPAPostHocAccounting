"""
/***************************************************************************
 MPAPostHocAccounting
                                 A QGIS plugin
 This plugin checks how your MPAs meet your placement objectives
                              -------------------
        begin                : 2018-09-24
        git sha              : $Format:%H$
        copyright            : (C) 2018 by Jonah Sullivan
        email                : jonahsullivan79@gmail.com
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
"""

import os

from osgeo import ogr
from PyQt5.QtCore import QFileInfo
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (
    QAction,
    QFileDialog,
    QHeaderView,
    QTableWidgetItem,
    QTreeWidgetItem,
    QTreeWidgetItemIterator,
)
from qgis.core import (
    QgsCoordinateReferenceSystem,
    QgsCoordinateTransformContext,
    QgsDistanceArea,
)

# Import the code for the dialog
from .MPA_postHocAccounting_dialog_base import MPAPostHocAccountingDialogBase
from .MPA_postHocAccounting_dialog_targets import MPAPostHocAccountingDialogTargets

# Initialize Qt resources from file resources.py
from .resources import *
from .xlsx_style import XlsxStyler

# xlwt 'lime' and 'rose', the palette colours this report has always used
GREEN = "99CC00"
RED = "FF99CC"


def id_field_type(values):
    """Use an integer column when every identifier is a whole number."""
    try:
        for value in values:
            if value is not None:
                int(value)
    except (TypeError, ValueError):
        return ogr.OFTString
    return ogr.OFTInteger


def add_sheet(data_source, name, fields):
    """Create a worksheet. fields: [(header, ogr type)]; headers become row 1."""
    sheet = data_source.CreateLayer(name, geom_type=ogr.wkbNone)
    for header, field_type in fields:
        sheet.CreateField(ogr.FieldDefn(header, field_type))
    return sheet


def add_row(sheet, values):
    """Append a row, leaving None values as empty cells."""
    feature = ogr.Feature(sheet.GetLayerDefn())
    for i, value in enumerate(values):
        if value is not None:
            feature.SetField(i, value)
    sheet.CreateFeature(feature)


def header_widths(header_cells):
    """Column widths matching the original xlwt report (1/256-character units)."""
    return {
        i + 1: (len(header) + 4) * 367 / 256 for i, header in enumerate(header_cells)
    }


class MPAPostHocAccounting:
    """QGIS Plugin Implementation."""

    def __init__(self, iface):
        """Constructor.

        :param iface: An interface instance that will be passed to this class
            which provides the hook by which you can manipulate the QGIS
            application at run time.
        :type iface: QgsInterface
        """
        # Save reference to the QGIS interface
        self.iface = iface
        # initialize plugin directory
        self.plugin_dir = os.path.dirname(__file__)

        # Declare instance attributes
        self.actions = []
        self.menu = "MPA Results Analysis"
        self.toolbar = self.iface.addToolBar("MPAPostHocAccounting")
        self.toolbar.setObjectName("MPAPostHocAccounting")

        # initialise and clear variables
        self.out_xlsx = None
        self.in_mpa_field = None
        self.in_map_layer = None
        self.check_poly_dict = None
        self.dlg_base = None
        self.dlg_targets = None

    def add_action(
        self,
        icon_path,
        text,
        callback,
        enabled_flag=True,
        add_to_menu=True,
        add_to_toolbar=True,
        status_tip=None,
        whats_this=None,
        parent=None,
    ):

        # Create the dialog and keep reference
        self.dlg_base = MPAPostHocAccountingDialogBase()
        self.dlg_targets = MPAPostHocAccountingDialogTargets()

        icon = QIcon(icon_path)
        action = QAction(icon, text, parent)
        action.triggered.connect(callback)
        action.setEnabled(enabled_flag)

        if status_tip is not None:
            action.setStatusTip(status_tip)

        if whats_this is not None:
            action.setWhatsThis(whats_this)

        if add_to_toolbar:
            self.toolbar.addAction(action)

        if add_to_menu:
            self.iface.addPluginToMenu(self.menu, action)

        self.actions.append(action)

        return action

    def initGui(self):
        """Create the menu entries and toolbar icons inside the QGIS GUI."""
        icon_path = ":/plugins/MPAPostHocAccounting/icon.png"
        self.add_action(
            icon_path,
            text="MPA Results Analysis",
            callback=self.run,
            parent=self.iface.mainWindow(),
        )

    def unload(self):
        """Removes the plugin menu item and icon from QGIS GUI."""
        for action in self.actions:
            self.iface.removePluginMenu("&MPA Results Analysis", action)
            self.iface.removeToolBarIcon(action)
        # remove the toolbar
        del self.toolbar

    def run(self):
        """Run method that performs all the real work"""

        # clear old variables
        self.dlg_base.fieldComboBox.setLayer(None)
        # self.dlg_base.inMPA_Layer.clear()
        # self.dlg_base.fieldComboBox.clear()
        iterator = QTreeWidgetItemIterator(
            self.dlg_base.inData, QTreeWidgetItemIterator.All
        )
        while iterator.value():
            iterator.value().takeChildren()
            iterator += 1
        i = self.dlg_base.inData.topLevelItemCount()
        while i > -1:
            self.dlg_base.inData.takeTopLevelItem(i)
            i -= 1

        # show the window
        self.dlg_base.show()

        # select the MPA layer
        self.in_map_layer = self.dlg_base.inMPA_Layer.currentLayer()

        def set_layer_name():
            self.in_map_layer = self.dlg_base.inMPA_Layer.currentLayer()

        self.dlg_base.inMPA_Layer.layerChanged.connect(set_layer_name)

        # set the mpaLayer for the field combo box
        def set_field_combo_box_layer(in_layer):
            self.dlg_base.fieldComboBox.setLayer(in_layer)

        self.dlg_base.inMPA_Layer.layerChanged.connect(set_field_combo_box_layer)

        # set the MPA unique identifier field
        def set_mpa_field():
            self.in_mpa_field = self.dlg_base.fieldComboBox.currentField()

        self.dlg_base.fieldComboBox.fieldChanged.connect(set_mpa_field)
        self.in_mpa_field = self.dlg_base.fieldComboBox.currentField()

        # add polygon layers and field names to tree widget
        def set_layers():
            # add layer names and field names to analysis selection window
            layer_fields_tree = self.dlg_base.inData
            layer_fields_tree.clear()
            for map_layer in self.iface.mapCanvas().layers():
                if map_layer.name() == self.in_map_layer.name():
                    pass
                else:
                    tree_item = QTreeWidgetItem()
                    layer_fields_tree.addTopLevelItem(tree_item)
                    tree_item.setText(0, map_layer.name())
                    for layer_field in map_layer.fields():
                        field_item = QTreeWidgetItem(tree_item)
                        field_item.setText(0, layer_field.name())
            self.dlg_base.inData.expandAll()

        self.dlg_base.fieldComboBox.fieldChanged.connect(set_layers)

        # add selected layers and fields to processing list
        self.check_poly_dict = {}

        def tree_selection_changed():
            self.check_poly_dict = {}
            get_selected = self.dlg_base.inData.selectedItems()
            for selected_item in get_selected:
                if selected_item.parent():
                    field_name = selected_item.text(0)
                    selected_layer_name = selected_item.parent().text(0)
                    for j in range(self.iface.mapCanvas().layerCount()):
                        selected_layer = self.iface.mapCanvas().layer(j)
                        if selected_layer.name() == selected_layer_name:
                            for layer_field in selected_layer.fields():
                                if layer_field.name() == field_name:
                                    self.check_poly_dict[selected_layer_name] = {
                                        "layer": selected_layer,
                                        "field": layer_field,
                                    }

        self.dlg_base.inData.itemSelectionChanged.connect(tree_selection_changed)

        # function returns dict of dicts with area of intersection for two shapefiles
        def intersect_area(layer1, field1, layer2, field2):
            area_dict = {}
            # loop through features in first shapefile
            for feat1 in layer1.getFeatures():
                feat_dict = {}
                geom1 = feat1.geometry()
                attr1 = feat1[layer1.fields().lookupField(field1)]
                # loop through features in second shapefile
                for feat2 in layer2.getFeatures():
                    geom2 = feat2.geometry()
                    attr2 = feat2[layer2.fields().lookupField(field2)]
                    # if features intersect then write feature attr and area to shape2 dict
                    if geom2.intersects(geom1):
                        intersection = geom1.intersection(geom2)
                        int_area = intersection.area()
                        if attr2 in feat_dict:
                            feat_dict[attr2] += int_area / geom1.area()
                        else:
                            feat_dict[attr2] = int_area / geom1.area()
                    # write shape2 dict to output dict
                area_dict[attr1] = feat_dict
            return area_dict

        # this part is executed after the ok button is pressed on the base window
        result_base = self.dlg_base.exec_()
        if result_base:
            table_widget = self.dlg_targets.tableWidget
            table_widget.setRowCount(0)
            row = 0
            for layer in self.check_poly_dict:
                table_widget.insertRow(row)
                table_item = QTableWidgetItem()
                table_item.setText(layer)
                table_widget.setItem(row, 0, table_item)
                row += 1
            for row in range(table_widget.rowCount()):
                for column in range(1, table_widget.columnCount()):
                    val = 0
                    if column == 1:
                        val = 10
                    elif column == 2:
                        val = 2
                    table_item = QTableWidgetItem()
                    table_item.setText(str(val))
                    table_widget.setItem(row, column, table_item)

            # resize columns to fit contents
            header = table_widget.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.Stretch)
            header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(2, QHeaderView.ResizeToContents)

            # clear the output dialog box before showing the
            self.dlg_targets.outTable.clear()

            # show the targets dialog box
            self.dlg_targets.show()

            # display file dialog to select output spreadsheet
            def out_file():
                out_name, _ = QFileDialog.getSaveFileName(
                    None,
                    "Output Spreadsheet",
                    os.getenv("HOME"),
                    "Spreadsheets (*.xlsx)",
                )
                out_path = QFileInfo(out_name).absoluteFilePath()
                if not out_path.upper().endswith(".XLSX"):
                    out_path = out_path + ".xlsx"
                if out_name:
                    self.out_xlsx = out_path
                    self.dlg_targets.outTable.setText(out_path)

            # select output spreadsheet file when browse button is clicked
            self.dlg_targets.outButton.clicked.connect(out_file)

            # this part is executed after the ok button is pressed on the targets window
            result_target = self.dlg_targets.exec_()
            if result_target:
                # set the coverage and replication targets
                for row in range(table_widget.rowCount()):
                    layer_name = ""
                    coverage_target = 0
                    repl_target = 0
                    for column in range(table_widget.columnCount()):
                        table_item = table_widget.item(row, column)
                        if column == 0:
                            layer_name = table_item.text()
                        elif column == 1:
                            coverage_target = table_item.text()
                        elif column == 2:
                            repl_target = table_item.text()
                    self.check_poly_dict[layer_name]["coverageTarget"] = int(
                        coverage_target
                    )
                    self.check_poly_dict[layer_name]["replTarget"] = int(repl_target)

                # GDAL writes the data; formatting is added once the file is closed
                if os.path.exists(self.out_xlsx):
                    os.remove(self.out_xlsx)
                data_source = ogr.GetDriverByName("XLSX").CreateDataSource(
                    self.out_xlsx
                )
                sheet_count = 0
                # {sheet number: ({cell reference: style key}, {column: width})}
                sheet_formats = {}

                # analyse mpa size and distance to nearest MPA
                # create dict of distances to other mpas
                if self.in_map_layer.featureCount() > 1:
                    dist_dict = {}
                    for feat in self.in_map_layer.getFeatures():
                        geom = feat.geometry()
                        attr = feat.attribute(self.in_mpa_field)
                        dist_list = []
                        attr_list = []
                        for test_feature in self.in_map_layer.getFeatures():
                            dist = geom.distance(test_feature.geometry())
                            if dist == 0.0:
                                pass
                            else:
                                # vertex in test feature closest to centroid of feature
                                closest_vertex = geom.closestVertex(
                                    test_feature.geometry().centroid().asPoint()
                                )
                                # vertex in feature closest to centroid of test feature
                                closest_vertex_test = (
                                    test_feature.geometry().closestVertex(
                                        geom.centroid().asPoint()
                                    )
                                )
                                # tool to measure distance between points (in metres)
                                d = QgsDistanceArea()
                                d.setEllipsoid("WGS84")
                                canvas_auth_id = (
                                    self.iface.mapCanvas()
                                    .mapSettings()
                                    .destinationCrs()
                                    .authid()
                                )
                                canvas_crs = QgsCoordinateReferenceSystem(
                                    canvas_auth_id
                                )
                                ellipsoid_crs = QgsCoordinateReferenceSystem(4326)
                                trans_context = QgsCoordinateTransformContext()
                                trans_context.calculateDatumTransforms(
                                    canvas_crs, ellipsoid_crs
                                )
                                d.setSourceCrs(canvas_crs, trans_context)
                                dist = d.measureLine(
                                    closest_vertex[0], closest_vertex_test[0]
                                )  # distance in metres
                                dist_list.append(dist)
                                attr_list.append(
                                    test_feature.attribute(self.in_mpa_field)
                                )

                        # add values to dictionary
                        mindist = min(dist_list)
                        minattr = attr_list[dist_list.index(mindist)]
                        dist_dict[attr] = [minattr, mindist]

                    # write closest mpa to workbook
                    header_cells = ["MPA ID", "Nearest MPA", "Distance (km)"]
                    id_type = id_field_type(dist_dict.keys())
                    sheet = add_sheet(
                        data_source,
                        "Distances",
                        [
                            (header_cells[0], id_type),
                            (header_cells[1], id_type),
                            (header_cells[2], ogr.OFTReal),
                        ],
                    )
                    sheet_count += 1
                    for item, (nearest, distance) in dist_dict.items():
                        if id_type == ogr.OFTInteger:
                            item = None if item is None else int(item)
                            nearest = None if nearest is None else int(nearest)
                        add_row(
                            sheet, [item, nearest, distance / 1000]
                        )  # conversion from metres to kilometres
                    sheet_formats[sheet_count] = ({}, header_widths(header_cells))

                # loop through polygon layers
                name_list = []
                for polyName in self.check_poly_dict:
                    # add a new worksheet to workbook, length limit is 31 characters
                    short_name = polyName[:30]
                    if short_name in name_list:
                        short_name = short_name[:29] + "1"
                    name_list.append(short_name)
                    # get information from the dictionary
                    layer = self.check_poly_dict[polyName]["layer"]
                    field = self.check_poly_dict[polyName]["field"]
                    coverage_target = self.check_poly_dict[polyName]["coverageTarget"]
                    repl_target = self.check_poly_dict[polyName]["replTarget"]
                    header_cells = [
                        polyName + " " + field.name(),
                        "Coverage" + " target=" + f"{coverage_target:.0f}%",
                        "Replication" + " target=" + str(repl_target),
                    ]
                    # create list of unique IDs for polygons
                    attr_index = layer.fields().lookupField(field.name())
                    attr_list = layer.uniqueValues(attr_index)
                    attr_list = list(attr_list)
                    attr_list.sort()
                    id_type = id_field_type(attr_list)
                    sheet = add_sheet(
                        data_source,
                        short_name,
                        [
                            (header_cells[0], id_type),
                            (header_cells[1], ogr.OFTReal),
                            (header_cells[2], ogr.OFTInteger),
                        ],
                    )
                    sheet_count += 1
                    # create dictionary with entry for each polygon with values of area intersecting with each MPA
                    mpa_area_per_poly = intersect_area(
                        layer, field.name(), self.in_map_layer, self.in_mpa_field
                    )
                    # print report, colouring each cell by whether it met its target
                    fills = {}
                    for row, uniqueID in enumerate(attr_list, start=2):
                        sum_area = sum(mpa_area_per_poly[uniqueID].values())
                        mpa_count = len(mpa_area_per_poly[uniqueID])
                        identifier = uniqueID
                        if id_type == ogr.OFTInteger and identifier is not None:
                            identifier = int(identifier)
                        add_row(sheet, [identifier, float(sum_area), mpa_count])
                        fills[f"B{row}"] = (
                            "coverage_met"
                            if sum_area >= coverage_target / 100.0
                            else "coverage_missed"
                        )
                        fills[f"C{row}"] = (
                            "replication_met"
                            if mpa_count >= repl_target
                            else "replication_missed"
                        )
                    sheet_formats[sheet_count] = (fills, header_widths(header_cells))

                # close the data source so the archive is complete, then style it
                data_source = None
                styler = XlsxStyler(self.out_xlsx)
                styles = {
                    "coverage_met": styler.add_style(fill=GREEN, num_fmt="0%"),
                    "coverage_missed": styler.add_style(fill=RED, num_fmt="0%"),
                    "replication_met": styler.add_style(fill=GREEN),
                    "replication_missed": styler.add_style(fill=RED),
                }
                for index, (fills, widths) in sheet_formats.items():
                    for ref, style_key in fills.items():
                        styler.set_cell(index, ref, styles[style_key])
                    styler.set_widths(index, widths)
                styler.save()
                if os.path.exists(self.out_xlsx):
                    try:
                        from os import startfile  # windows only

                        os.startfile(self.out_xlsx)
                    except ImportError:
                        import subprocess

                        subprocess.run(["xdg-open", self.out_xlsx])  # if not windows

            self.dlg_targets.outButton.clicked.disconnect(out_file)
