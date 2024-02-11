# -*- coding: utf-8 -*-
"""
/***************************************************************************
Name			 	 : pyxelSync 
Description          : Use pyxelSync
Date                 : 11/Feb/24
copyright            : (C) 2024 by J.D.M.P.E
email                : 
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
 This script initializes the plugin, making it known to QGIS.
"""

def classFactory(iface):
    # loads pyxelSync Class from PyXel_Sync Library
    from .PyXel_Sync import pyxelSync
    return pyxelSync(iface)


