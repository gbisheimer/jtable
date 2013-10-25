/************************************************************************
 * FORM TEMPLATES extension for jTable                                   *
 * Author: Guillermo Bisheimer                                           *
 *************************************************************************/
(function($) {

    //Reference to base object members
    var base = {
        _create: $.hik.jtable.prototype._create,
        _showEditForm: $.hik.jtable.prototype._showEditForm,
        _showAddRecordForm: $.hik.jtable.prototype._showAddRecordForm
    };

    //extension members
    $.extend(true, $.hik.jtable.prototype, {
        /************************************************************************
         * DEFAULT OPTIONS / EVENTS                                              *
         *************************************************************************/
        options: {
            dialogs: {}
        },
        /************************************************************************
         * OVERRIDED METHODS                                                     *
         *************************************************************************/

        /* Overrides base method to do editing-specific constructions.
         *************************************************************************/
        _create: function() {
            base._create.apply(this, arguments);

            if (this.options.dialogs !== undefined) {
                if (this.options.dialogs.template !== undefined) {
                    this.options.dialogs.template.remove();
                }
            }
        },
        /* Shows edit form for a row.
         *************************************************************************/
        _showEditForm: function($tableRow) {
            var self = this;

            base._showEditForm.apply(this, arguments);

            if (this.options.dialogs !== undefined) {
                var record = $tableRow.data('record');
                var $editForm = this._$editDiv.find('form#jtable-edit-form');
                var $dialog = this._$editDiv;
                var options = this.options.dialogs;

                this._applyDialogOptions({
                    dialog: $dialog,
                    form: $editForm,
                    formType: 'edit',
                    options: options,
                    record: record
                });
            }
        },
        /* Shows create form for a row.
         *************************************************************************/
        _showAddRecordForm: function() {
            var self = this;

            base._showAddRecordForm.apply(this, arguments);

            if (this.options.dialogs !== undefined) {
                var $editForm = this._$addRecordDiv.find('form#jtable-create-form');
                var $dialog = this._$addRecordDiv;
                var options = this.options.dialogs;

                this._applyDialogOptions({
                    dialog: $dialog,
                    form: $editForm,
                    formType: 'create',
                    options: options,
                    record: null
                });
            }
        },
        /************************************************************************
         * PRIVATE METHODS                                                       *
         *************************************************************************/

        /* Apply custom options to forms
         *************************************************************************/
        _applyDialogOptions: function(data) {
            var self = this;

            /*dialog: $dialog,
             form: $editForm,
             formType: 'edit',
             options: options,
             record: record*/

            // Apply user defined dialog options
            if (data.options.options) {
                data.dialog.dialog('option', data.options.options );
            }

            // Apply aditional field options
            $.each(data.form.find('div.jtable-input-field-container'), function(index, fieldContainer) {
                var $fieldContainer = $(fieldContainer);
                var $field = $fieldContainer.find('div.jtable-input').children();
                var fieldName = $field.attr('name');
                var fieldOptions = self.options.fields[fieldName];
                if (fieldOptions.readOnly === true ||
                        (typeof fieldOptions.readOnly === 'object' && fieldOptions.readOnly[data.formType] === true)) {
                    $field.attr('disabled', 'disabled');
                }
            });

            if (data.options.template) {
                //Move created fields to template
                $.each(data.form.find('div.jtable-input-field-container'), function(index, fieldContainer) {
                    var $fieldContainer = $(fieldContainer);
                    var fieldName = $fieldContainer.find('div.jtable-input').children().attr('name');
                    var $templateFieldDiv = data.options.template.find('div#' + fieldName);

                    // Checks if field is included in template
                    if ($templateFieldDiv.length !== 0) {
                        $templateFieldDiv
                                .empty()
                                .append($fieldContainer);
                    }
                });
                
                //Adds template to form
                data.form.append($.extend({}, data.options.template));                
            }

            // Dialog button
            if (data.options.buttons) {
                $.each(data.options.buttons, function(childID, childOptions) {
                    var $childDiv = data.form.find('div#' + childID);
                    var params = {
                        form: data.form,
                        record: data.record
                    };
                    
                    // Get child options
                    //var options = typeof(childOptions.options) === 'function' ? childOptions.options(params) : childOptions.options || {};

                    // if not present in template already, appends element to edit Form
                    if ($childDiv.length === 0) {
                        var dialogButtons = data.dialog.dialog('option', 'buttons');
                        var index;
                        // Checks if buttons already exists in Form
                        for (index = 0; index < dialogButtons.length; index++) {
                            if (dialogButtons[index].id === childID)
                                break;
                        }
                        // Adds button to Form if new
                        if (index === dialogButtons.length) {
                            dialogButtons.splice(0, 0, {
                                id: childID,
                                text: childOptions.label || childID,
                                click: function(e) {
                                    e.preventDefault();
                                    if (childOptions.click) {
                                        childOptions.click(params);
                                    }
                                }
                            });
                            data.dialog.dialog('option', 'buttons', dialogButtons);
                        }
                    }

                    if ($childDiv.length !== 0) {
                        $childDiv.button(childOptions)
                                .click(function(e) {
                            if (childOptions.click) {
                                childOptions.click(params);
                            }
                        });
                    }
                });
            }

            data.dialog.dialog("option", "position", {
                my: "center",
                at: "center"
            });
        }
    });

})(jQuery);