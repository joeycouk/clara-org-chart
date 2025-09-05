(ns clara-org-chart.org-chart-builder
(:require
   [clara-xlsx-diff.util :as util]
   [clara-xlsx-diff.xlsx :as xlsx]
   [clara-xlsx-diff.eav :as eav]
   #?@(:clj [[clara.rules :as rules]
             [clara-eav.rules :as er]
             [clara.rules.accumulators :as accum]
            ;;  [portal.api :as p] ;; must comment out before building cljs])
))