/************************************************************************
 * LOCAL DB  extension for jTable
 * Author: Guillermo Bisheimer
 * Version 1.0
 *************************************************************************/
(function($) {

    //Reference to base object members
    var base = {
        _create: $.hik.jtable.prototype._create,
        _ajax: $.hik.jtable.prototype._ajax
    };

    //extension members
    $.extend(true, $.hik.jtable.prototype, {
        /************************************************************************
         * DEFAULT OPTIONS / EVENTS                                              *
         *************************************************************************/
        options: {
            DB: [],
            
            //clientOnly: true,
            //delayCommit: true,
            
            // Events
            onDBList: function (event, data) { },
            onDBCreate: function (event, data) { },
            onDBUpdate: function (event, data) { },
            onDBDelete: function (event, data) { },
                        
            // Messages
            messages: {
                localDB: {
                    noKeyField: 'Primary key not defined. Cannot update table data.',
                    recordNotFound: 'No matching record found.',
                    multipleRecordsFound: 'Multiple records found. Check primary key field definition.'
                }
            }
        },
        /************************************************************************
         * OVERRIDED METHODS                                                     *
         *************************************************************************/

        /* Overrides base method to do editing-specific constructions.
         *************************************************************************/
        _create: function() {
            var self = this;
            base._create.apply(this, arguments);

            if (this.options.clientOnly === true) {
                $.each(this.options.actions, function(index, action) {
                    self.options.actions[index + 'Backup'] = action;
                    self.options.actions[index] = index;
                });
            }
        },
        /* Overrides base method
         *************************************************************************/
        _ajax: function(options) {
            var self = this;

            // Gets which action is calling the ajax function
            var action = Object.keys(self.options.actions).filter(function(key) {
                return options.url.indexOf(self.options.actions[key]) === 0;
            });

            // Check if we are working with client side DB
            if (this.options.clientOnly !== true ||
                    action.length === 0) {
                return base._ajax.apply(this, arguments);
            }

            var opts = $.extend({}, this.options.ajaxSettings, options);

            // Override success
            opts.success = function(data) {
                if (options.success) {
                    options.success(data);
                }
            };

            // Override error
            opts.error = function() {
                if (options.error) {
                    options.error();
                }
            };

            // Override complete
            opts.complete = function() {
                if (options.complete) {
                    options.complete();
                }
            };

            // Get parameters
            var paramString = '';
            if (options.url.indexOf('?') >= 0)
                paramString += options.url.split('?')[1];
            if (typeof options.data === 'string')
                paramString += options.data;
            if (typeof options.data === 'object')
                paramString += $.param(options.data);
            var params = paramString ? JSON.parse('{"' + paramString.replace(/&/g, '","').replace(/=/g, '":"') + '"}', function(key, value) {
                return key === "" ? value : decodeURIComponent(value);
            }) : {};

            var callbacks = {
                'listAction': function(data) {
                    return self._onDBList('list', data);
                },
                'updateAction': function(data) {
                    return self._onDBUpdate('update', data);
                },
                'createAction': function(data) {
                    return self._onDBCreate('create', data);
                },
                'deleteAction': function(data) {
                    return self._onDBDelete('delete', data);
                }
            };

            /* TODO
             * - Agregar registro de eventos, para poder reenviar los datos modificados a la base de datos
             */
            var result = callbacks[action](params);

            if (result !== false) {
                opts.success(result);
            }
            else {
                opts.error();
            }

            opts.complete();
        },
        /**
         * Adds record to local database
         * @param {object} record Record to store
         * @returns {undefined}
         */
        _addRecordToLocalDB: function(record) {
            this.options.DB.push(record);
        },
        /**
         * Updates record in local database
         * @param {object} data Data to update in target record
         * @param {object} record Target record to update
         * @returns {undefined}
         */
        _updateRecordInLocalDB: function(data, record) {
            // Adds field values to record
            for (var field in data) {
                record[field] = data[field];
            }
        },
        /**
         * Deletes record from local database
         * @param {integer} recordIndex
         * @returns {undefined}
         */
        _deleteRecordFromLocalDB: function(recordIndex) {
            this.options.DB.splice(recordIndex, 1);
        },
        /**
         * Local DB listAction callback
         * @param {string} event Event type ('list', 'create', 'update', 'delete')
         * @param {object} data Object posted parameters
         * @returns {object} result object acording to jTable server-side specifications
         */
        _onDBList: function(event, data) {
            var self = this;

            /* TODO
             * - Ordenamiento de datos
             * - Paginación
             */
            self._trigger("onDBList", 'list', {DB: self.options.DB});

            var result = {
                Result: 'OK',
                TotalRecordCount: self.options.DB.length,
                Records: self.options.DB
            };

            return result;
        },
        /**
         * Local DB createAction callback
         * @param {string} event 'create'
         * @param {object} data Object posted parameters
         * @returns {object} result object acording to jTable server-side specifications
         */
        _onDBCreate: function(event, data) {
            var self = this;

            var record = {};
            var result = {};
            var success = true;
            var message = '';

            // Adds table fields to record
            for (var index in self._fieldList) {
                var fieldName = self._fieldList[index];
                var fieldOptions = self.options.fields[fieldName];
                if (!fieldOptions.childTable) {
                    record[fieldName] = fieldOptions.defaultValue || null;
                }
            }

            // Adds field values to record
            for (var field in data) {
                record[field] = data[field];
            }

            success = self._trigger("onDBCreate", 'create', {data: data, record: record, DB: self.options.DB, message: message});

            if (success) {
                result = {
                    Result: 'OK',
                    Record: record
                };

                // Stores new record in local database
                this._addRecordToLocalDB(record);
            }
            else {
                result = {
                    Result: 'ERROR',
                    Message: message
                };
            }

            return result;
        },
        /**
         * Local DB updateAction callback
         * @param {string} event 'update'
         * @param {object} data Object posted parameters
         * @returns {object} result object acording to jTable server-side specifications
         */
        _onDBUpdate: function(event, data) {
            var self = this;

            var success = true;
            var message = '';
            var keyField = self._keyField;
            var recordIndex = null;
            var result = {
                Result: 'ERROR',
                Message: message
            };

            if (!keyField || !data.hasOwnProperty(keyField)) {
                result.Message = self.options.messages.localDB.noKeyField;
                return result;
            }

            // Gets target record
            var records = self.options.DB.filter(function(record, index) {
                if (record[keyField] == data[keyField])
                {
                    recordIndex = index;
                    return true;
                }
            });

            // If there no coincidence on key field value, returns error
            if (records.length === 0) {
                result.Message = self.options.messages.localDB.recordNotFound;
                return result;
            }

            // If there is more than one coincidence on key field value, returns error
            if (records.length === 0) {
                result.Message = self.options.messages.localDB.multipleRecordsFound;
                return result;
            }

            var record = records[0];

            success = self._trigger("onDBUpdate", 'update', {recordIndex: recordIndex, keyField: keyField, data: data, record: record, DB: self.options.DB, message: message});

            if (success) {
                result = {
                    Result: 'OK',
                    Record: record
                };

                // Stores updated record in local database
                this._updateRecordInLocalDB(data, record);
            }
            else {
                result = {
                    Result: 'ERROR',
                    Message: message
                };
            }

            return result;
        },
        /**
         * Local DB deleteAction callback
         * @param {string} event 'delete'
         * @param {object} data Object posted parameters
         * @returns {object} result object acording to jTable server-side specifications
         */
        _onDBDelete: function(event, data) {
            var self = this;

            var success = true;
            var message = '';
            var keyField = self._keyField;
            var recordIndex = null;
            var result = {
                Result: 'ERROR',
                Message: message
            };

            if (!keyField || !data.hasOwnProperty(keyField)) {
                result.Message = self.options.messages.localDB.noKeyField;
                return result;
            }

            // Gets target record
            var records = self.options.DB.filter(function(record, index) {
                if (record[keyField] == data[keyField])
                {
                    recordIndex = index;
                    return true;
                }
            });

            // If there no coincidence on key field value, returns error
            if (records.length === 0) {
                result.Message = self.options.messages.localDB.recordNotFound;
                return result;
            }

            // If there is more than one coincidence on key field value, returns error
            if (records.length === 0) {
                result.Message = self.options.messages.localDB.multipleRecordsFound;
                return result;
            }

            var record = records[0];

            success = self._trigger("onDBDelete", 'delete', {recordIndex: recordIndex, keyField: keyField, data: data, record: record, DB: self.options.DB, message: message});

            if (success) {
                result = {
                    Result: 'OK',
                    Record: record
                };

                // Deletes record from local database
                this._deleteRecordFromLocalDB(recordIndex);
            }
            else {
                result = {
                    Result: 'ERROR',
                    Message: message
                };
            }
            return result;
        },
        
        /************************************************************************
        * PUBLIC METHODS                                                        *
        *************************************************************************/

        /**
         * Gets records stored on local DB
         * @returns {object} Local database object
         */
        getTableRecords: function () {
            return this.options.DB;
        },
        /**
         * Clear records in DB
         * @returns {undefined}
         */
        clearTableRecords: function () {
            this.options.DB.length = 0;
        }
    });

})(jQuery);