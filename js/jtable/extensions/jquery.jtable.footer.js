/************************************************************************
* FOOTER extension for jTable                                           *
* Author: Guillermo Bisheimer                                           
* Rev. 1.0
*************************************************************************/
(function($) {

    //Reference to base object members
    var base = {
        _create: $.hik.jtable.prototype._create,
        _onRowsRemoved: $.hik.jtable.prototype._onRowsRemoved,
        _onRecordsLoaded: $.hik.jtable.prototype._onRecordsLoaded,
        _onRecordAdded: $.hik.jtable.prototype._onRecordAdded,
        _onRecordUpdated: $.hik.jtable.prototype._onRecordUpdated
    };

    //extension members
    $.extend(true, $.hik.jtable.prototype, {

        /************************************************************************
        * DEFAULT OPTIONS / EVENTS                                              *
         *************************************************************************/
        options: {
            footer: false
        },

        /************************************************************************
        * PRIVATE FIELDS                                                        *
        *************************************************************************/

        _$tfoot: null, //Reference to the footer area in bottom panel

        /************************************************************************
        * OVERRIDED METHODS                                                     *
        *************************************************************************/

        /* Overrides base method to create footer constructions.
         *************************************************************************/
        _create: function() {
            base._create.apply(this, arguments);
            if( this.options.footer ){
                this._createTableFoot();
            }
        },
        
        /* Overrides _onRecordAdded method to re-load table when a new row is created.
        *************************************************************************/
        _onRecordAdded: function(data) {
            if (this.options.footer) {
                this._updateTableFoot();
            }
            base._onRecordAdded.apply(this, arguments);
        },

        /* Overrides _onRecordUpdated method to re-load table when a new row is created.
        *************************************************************************/
        _onRecordUpdated: function($row, options) {
            if (this.options.footer) {
                this._updateTableFoot();
            }
            base._onRecordUpdated.apply(this, arguments);
        },

        /* Overrides _onRowsRemoved method to re-load table when a row is removed from table.
        *************************************************************************/
        _onRowsRemoved: function($rows, reason) {
            if (this.options.footer && reason != 'reloading') {
                this._updateTableFoot();
            }
            base._onRowsRemoved.apply(this, arguments);
        },

        /* Overrides _onRecordsLoaded method to to render footer row.
        *************************************************************************/
        _onRecordsLoaded: function(data) {
            if( this.options.footer ){
                this._updateTableFoot( data );
            }
            base._onRecordsLoaded.apply(this, arguments);
        },
        
        /************************************************************************
        * PRIVATE METHODS                                                       *
        *************************************************************************/
        _updateTableFoot:  function ( data ) {
            var self = this;
            
            // If no data was provided, retrieve data from table rows
            if( data === undefined ){
                data = {Records:[]};
                $.each(this._$tableRows, function(index, row) {
                    data.Records.push(row.data('record'));
                });
            }
            
            this._$tfoot.find('th').each(function (index, cell) {
                var $cell = $(cell);
                var fieldName = $cell.data('fieldName')
                if( fieldName && self.options.fields[fieldName].footer )
                {   
                    $cell.find('span')
                    .empty()
                    .append(self.options.fields[fieldName].footer(data));
                }
            });
        },

        /* Creates footer (all column footers) of the table.
        *************************************************************************/
        _createTableFoot: function () {
            this._$tfoot = $('<tfoot></tfoot>').appendTo(this._$table);
            this._addRowToTableFoot(this._$tfoot);
        },

        /* Adds tr element to given tfoot element
        *************************************************************************/
        _addRowToTableFoot: function ($tfoot) {
            var $tr = $('<tr></tr>').appendTo($tfoot);
            this._addColumnsToFooterRow($tr);
        },

        /* Adds column footer cells to given tr element.
        *************************************************************************/
        _addColumnsToFooterRow: function ($tr) {
            for (var i = 0; i < this._columnList.length; i++) {
                var fieldName = this._columnList[i];
                var $footerCell = this._createFooterCellForField(fieldName, this.options.fields[fieldName]);
                $footerCell.data('fieldName', fieldName).appendTo($tr);
            }
            
            if (this.options.actions.updateAction !== undefined) {
                this._createFooterCellForField(null, {})
                .appendTo($tr);
            }
            
            if (this.options.actions.deleteAction !== undefined) {
                this._createFooterCellForField(null, {})
                .appendTo($tr);
            }
        },

        /* Creates a header cell for given field.
        *  Returns th jQuery object.
        *************************************************************************/
        _createFooterCellForField: function (fieldName, field) {
            return $('<th class="jtable-column-footer">' +
                '<div class="jtable-column-footer-container"><span class="jtable-column-footer-text"></span></div></th>');
        }
    });

})(jQuery);