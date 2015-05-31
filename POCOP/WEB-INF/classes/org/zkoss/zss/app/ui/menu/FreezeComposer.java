package org.zkoss.zss.app.ui.menu;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.ui.Spreadsheet;

public class FreezeComposer extends SelectorComposer<Component> {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	@Wire
    private Spreadsheet ss;
 
    @Listen("onClick = #freezeButton")
    public void freeze() {
        Ranges.range(ss.getSelectedSheet())
        .setFreezePanel(ss.getSelection().getRow(), ss.getSelection().getColumn());
    }
     
    @Listen("onClick = #unfreezeButton")
    public void unfreeze() {
        Ranges.range(ss.getSelectedSheet()).setFreezePanel(0,0);
    }
}
