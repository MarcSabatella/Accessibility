
__id__        = "$Id$"
__version__   = "$Revision$"
__date__      = "$Date$"
__copyright__ = "Copyright (c) 2020 Marc Sabatella"
__license__   = "LGPL"

import pyatspi

import orca.debug as debug
import orca.orca as orca
import orca.orca_state as orca_state
import orca.scripts.default as default
import orca.messages as messages

from orca.scripts.toolkits import Qt

#orca.debug.debugLevel = orca.debug.LEVEL_ALL

class Script(Qt.Script):

    def __init__(self, app):
        super().__init__(app)
        msg = "Orca for MuseScore 3"
        self.presentMessage(msg)

    def onValueChanged(self, event):
        """Callback for object:property-change:accessible-value events."""
        obj = event.source
        role = obj.getRole()
        if role == pyatspi.ROLE_LABEL and obj == orca_state.locusOfFocus:
            msg = obj.description
            self.presentationInterrupt()
            self.presentMessage(msg)
        super().onValueChanged(event)

"""
    # TODO: attempt at supporting key signature selector in new score wizard
    def onNameChanged(self, event):
        obj = event.source
        if obj == orca_state.locusOfFocus:
            names = self.pointOfReference.get('names', {})
            oldName = names[hash(obj)]
            newName = obj.name
            if newName != oldName:
                self.presentMessage(newName)
        super().onNameChanged(event)
"""

"""
    # TODO: attempt at handling description changed events
    def onDescriptionChanged(self, event):
        msg = "description changed"
        self.presentMessage(msg)
        return
"""

"""
    # TODO: attempt at installing handler for description changed events
    def getListeners(self):
        listeners = super().getListeners()
        listeners["object:description-change:accessible-description"] = self.onDescriptionChanged
        return listeners
"""

"""
    # TODO: attempt at correcting Orca reading of spin button label vs. name
    def locusOfFocusChanged(self, event, oldLocusOfFocus, newLocusOfFocus):
        if newLocusOfFocus == oldLocusOfFocus:
            return
        if newLocusOfFocus.getRole() == pyatspi.ROLE_SPIN_BUTTON:
            msg = newLocusOfFocus.name
            self.presentMessage(msg)
            msg = newLocusOfFocus.queryValue().currentvalue
            self.presentMessage(msg)
        else:
            super().locusOfFocusChanged(event, oldLocusOfFocus, newLocusOfFocus)
"""
