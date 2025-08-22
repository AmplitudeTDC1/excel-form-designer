/* ============================================================================
 * bindingManager.js
 * ----------------------------------------------------------------------------
 * Purpose:
 *   Centralized manager for creating and handling Excel bindings
 *   between form controls and Excel ranges.
 *
 * Usage:
 *   1. Load this file before form_designer_panel.js in your add-in HTML.
 *   2. Use BindingManager.bindControlToRange(controlId, rangeName)
 *      to connect a form control with an Excel range.
 *   3. Binding is two-way:
 *        - Control → Excel (on user input/change)
 *        - Excel → Control (on Excel binding data change)
 *
 * Author: Sprint 2 Prototype
 * ============================================================================
 */

var BindingManager = (function () {
  // Private dictionary of bindings: { controlId: bindingId }
  const controlBindings = {};

  // Initialize event handlers for Excel binding data changes
  function attachBindingChangedHandler(binding) {
    binding.addHandlerAsync(
      Office.EventType.BindingDataChanged,
      function (eventArgs) {
        console.log("Excel → Form update:", eventArgs.binding.id);
        refreshControlFromExcel(eventArgs.binding.id);
      },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to attach BindingDataChanged handler:", result.error.message);
        }
      }
    );
  }

  // Push control value → Excel
  function pushToExcel(controlId, bindingId) {
    const element = document.getElementById(controlId);
    if (!element) return;

    let value = getControlValue(element);

    Office.context.document.bindings.getByIdAsync(bindingId, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        let binding = asyncResult.value;
        binding.setDataAsync(
          [[value]],
          { coercionType: Office.CoercionType.Matrix },
          function (res) {
            if (res.status === Office.AsyncResultStatus.Failed) {
              console.error("Error pushing to Excel:", res.error.message);
            } else {
              console.log(`Form → Excel updated: ${controlId} → ${bindingId}`);
            }
          }
        );
      }
    });
  }

  // Pull Excel value → Control
  function refreshControlFromExcel(bindingId) {
    Office.context.document.bindings.getByIdAsync(bindingId, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        let binding = asyncResult.value;
        binding.getDataAsync({ coercionType: Office.CoercionType.Matrix }, function (res) {
          if (res.status === Office.AsyncResultStatus.Succeeded) {
            let value = res.value[0][0]; // Top-left cell
            let controlId = getControlIdByBinding(bindingId);
            let element = document.getElementById(controlId);
            if (element) {
              setControlValue(element, value);
              console.log(`Excel → Form updated: ${bindingId} → ${controlId}`);
            }
          }
        });
      }
    });
  }

  // Utility: get controlId from bindingId
  function getControlIdByBinding(bindingId) {
    for (let controlId in controlBindings) {
      if (controlBindings[controlId] === bindingId) return controlId;
    }
    return null;
  }

  // Helpers to extract/set control values
  function getControlValue(element) {
    if (element.type === "checkbox") {
      return element.checked ? "TRUE" : "FALSE";
    } else if (element.tagName === "SELECT") {
      return element.value;
    } else {
      return element.value;
    }
  }

  function setControlValue(element, value) {
    if (element.type === "checkbox") {
      element.checked = value.toString().toLowerCase() === "true";
    } else {
      element.value = value;
    }
  }

  return {
    /**
     * Create or re-use a binding between a form control and Excel range
     * @param {string} controlId - ID of the form control (textbox, dropdown, etc.)
     * @param {string} rangeName - Named range or address in Excel (e.g., "Sheet1!B2")
     */
    bindControlToRange: function (controlId, rangeName) {
      const bindingId = `bind_${controlId}`;

      // Store mapping
      controlBindings[controlId] = bindingId;

      // Create binding
      Office.context.document.bindings.addFromNamedItemAsync(
        rangeName,
        Office.BindingType.Matrix,
        { id: bindingId },
        function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log(`Binding created: ${controlId} ↔ ${rangeName}`);
            attachBindingChangedHandler(result.value);

            // Sync initial value from Excel → control
            refreshControlFromExcel(bindingId);

            // Listen to control changes
            let element = document.getElementById(controlId);
            if (element) {
              element.addEventListener("change", function () {
                pushToExcel(controlId, bindingId);
              });
            }
          } else {
            console.error("Binding failed:", result.error.message);
          }
        }
      );
    },

    /**
     * Force a refresh from Excel → control
     * @param {string} controlId
     */
    refreshControl: function (controlId) {
      let bindingId = controlBindings[controlId];
      if (bindingId) {
        refreshControlFromExcel(bindingId);
      }
    },

    /**
     * Force a push from control → Excel
     * @param {string} controlId
     */
    updateExcelFromControl: function (controlId) {
      let bindingId = controlBindings[controlId];
      if (bindingId) {
        pushToExcel(controlId, bindingId);
      }
    }
  };
})();
