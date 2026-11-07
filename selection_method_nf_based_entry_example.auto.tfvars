####################
### Xpath: /managed-element/call-function/amf-function/network-function/selection-method/selection-method-nf-based/selection-method-nf-based-entry
### CLI Commands: CRTE-SEL-MTH:NFTYPE,SELMTH;
### Description:
### This list manages the entry of selection method based on NF.
### Dependencies: None
####################

selection_method_nf_based_entry = [
  {
    sel_mth_nftype     = "5geir"
    sel_mth_sel_method = "indscp"
    terrayang = {
      # with_defaults = "report-all"
    }
  },
  {
    sel_mth_nftype     = "ausf"
    sel_mth_sel_method = "indscp"
    terrayang = {
      # with_defaults = "report-all"
    }
  },
  {
    sel_mth_nftype     = "pcf"
    sel_mth_sel_method = "local"
    terrayang = {
      # with_defaults = "report-all"
    }
  },
  {
    sel_mth_nftype     = "smf"
    sel_mth_sel_method = "indscp"
    terrayang = {
      # with_defaults = "report-all"
    }
  },
  {
    sel_mth_nftype     = "smsf"
    sel_mth_sel_method = "indscp"
    terrayang = {
      # with_defaults = "report-all"
    }
  },
  {
    sel_mth_nftype     = "udm"
    sel_mth_sel_method = "indscp"
    terrayang = {
      # with_defaults = "report-all"
    }
  },
]
