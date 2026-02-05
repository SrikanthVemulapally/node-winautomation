{
  "targets": [
    {
      "target_name": "Automation",
      "cflags!": [ "-fno-exceptions" ],
      "cflags_cc!": [ "-fno-exceptions" ],
      "msvs_settings": {
        "VCCLCompilerTool": { "ExceptionHandling": 1 },
      },
      "sources": [ 
            # UI Automation Feature
            "features/ui-automation/src/AutomationAddon.cc",
            
            # COM Automation Feature
            "features/com-automation/src/core/COMLifecycle.cc",
            "features/com-automation/src/core/COMVariant.cc",
            "features/com-automation/src/core/COMDispatchWrapper.cc",
            "features/com-automation/src/core/COMEventSink.cc",
            "features/com-automation/src/core/COMObject.cc",

            # UI Automation - Enumerations
            "features/ui-automation/src/enumerations/AnnotationTypeIdsWrapper.cc",
            "features/ui-automation/src/enumerations/AttributeIdsWrapper.cc",
            "features/ui-automation/src/enumerations/ControlTypeIdsWrapper.cc",
            "features/ui-automation/src/enumerations/DockPositionsWrapper.cc",
            "features/ui-automation/src/enumerations/ElementModesWrapper.cc",
            "features/ui-automation/src/enumerations/EventIdsWrapper.cc",
            "features/ui-automation/src/enumerations/ExpandCollapseStatesWrapper.cc",
            "features/ui-automation/src/enumerations/OrientationTypesWrapper.cc",
            "features/ui-automation/src/enumerations/PatternIdsWrapper.cc",
            "features/ui-automation/src/enumerations/PropertyIdsWrapper.cc",
            "features/ui-automation/src/enumerations/ProviderOptionsWrapper.cc",
            "features/ui-automation/src/enumerations/RowOrColumnMajorWrapper.cc",
            "features/ui-automation/src/enumerations/StyleIdsWrapper.cc",
            "features/ui-automation/src/enumerations/ScrollAmountsWrapper.cc",
            "features/ui-automation/src/enumerations/SupportedTextSelectionsWrapper.cc",
            "features/ui-automation/src/enumerations/SynchronizedInputTypesWrapper.cc",
            "features/ui-automation/src/enumerations/TextPatternRangeEndpointWrapper.cc",
            "features/ui-automation/src/enumerations/TextUnitsWrapper.cc",
            "features/ui-automation/src/enumerations/ToggleStatesWrapper.cc",
            "features/ui-automation/src/enumerations/TreeScopesWrapper.cc",
            "features/ui-automation/src/enumerations/WindowInteractionStatesWrapper.cc",
            "features/ui-automation/src/enumerations/WindowVisualStatesWrapper.cc",
            "features/ui-automation/src/enumerations/ZoomUnitsWrapper.cc",

            # UI Automation - Utilities
            "features/ui-automation/src/utilities/AutomationEventHandler.cc",
            "features/ui-automation/src/utilities/FocusChangedEventHandler.cc",
            "features/ui-automation/src/utilities/PropertyChangedEventHandler.cc",
            "features/ui-automation/src/utilities/StructureChangedEventHandler.cc",
            "features/ui-automation/src/utilities/Functions.cc",

            # UI Automation - Wrappers
            "features/ui-automation/src/wrappers/RectWrapper.cc",
            "features/ui-automation/src/wrappers/IUnknownWrapper.cc",

            # Desktop Management Feature (temporarily disabled - needs Windows SDK)
            # "features/desktop-management/src/wrappers/ChildSessionWrapper.cc",
            # "features/desktop-management/src/wrappers/RDPClientWrapper.cc",
            # "features/desktop-management/src/wrappers/DesktopManagerWrapper.cc",

            "features/ui-automation/src/wrappers/IUIAutomationElementWrapper.cc" ,
            "features/ui-automation/src/wrappers/IUIAutomationWrapper.cc", 
            "features/ui-automation/src/wrappers/IUIAutomationConditionWrapper.cc",
            "features/ui-automation/src/wrappers/IUIAutomationElementArrayWrapper.cc",
            "features/ui-automation/src/wrappers/IUIAutomationTreeWalkerWrapper.cc",
            "features/ui-automation/src/wrappers/IUIAutomationCacheRequestWrapper.cc", 
            "features/ui-automation/src/wrappers/IUIAutomationProxyFactoryMappingWrapper.cc",
            "features/ui-automation/src/wrappers/IUIAutomationFocusChangedEventHandlerWrapper.cc",
            "features/ui-automation/src/wrappers/IUIAutomationEventHandlerWrapper.cc",
            "features/ui-automation/src/wrappers/IUIAutomationPropertyChangedEventHandlerWrapper.cc",
            "features/ui-automation/src/wrappers/IUIAutomationStructureChangedEventHandlerWrapper.cc",

            # UI Automation - Patterns
            "features/ui-automation/src/patterns/IUIAutomationTextRangeWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationAnnotationPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationDockPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationDragPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationDropTargetPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationExpandCollapsePatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationGridItemPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationGridPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationInvokePatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationItemContainerPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationLegacyIAccessiblePatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationMultipleViewPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationRangeValuePatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationScrollItemPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationScrollPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationSelectionItemPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationSelectionPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationSelectionPattern2Wrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationSpreadsheetItemPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationSpreadsheetPatternWrapper.cc", 
            "features/ui-automation/src/patterns/IUIAutomationStylesPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationSynchronizedInputPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationTableItemPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationTablePatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationTextChildPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationTextEditPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationTextPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationTextPattern2Wrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationTogglePatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationTransformPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationTransformPattern2Wrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationValuePatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationVirtualizedItemPatternWrapper.cc",
            "features/ui-automation/src/patterns/IUIAutomationWindowPatternWrapper.cc",
      ],
      "include_dirs": ["<!(node -p \"require('node-addon-api').include_dir\")"],
      "libraries" : [
        "uiautomationcore.lib",
        "Wtsapi32.lib",
        "Userenv.lib",
        "ole32.lib",
        "oleaut32.lib",
      ]
    }
  ]
}