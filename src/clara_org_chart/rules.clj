(ns clara-org-chart.rules
  (:require
             [clara-org-chart.xlsx :as xlsx]
             [clara-org-chart.eav :as eav]
             [clara-org-chart.position :as pos]
             [clara.rules :as rules]
             [clara.rules.accumulators :as accum])
  (:import (clara_org_chart.position Position))
  )


(rules/defquery get-position-values
  "Query to get all position values"
  []
  [?position-values <- (accum/all) :from [Position]])


(comment

  (def eavs (eav/xlsx->eav (xlsx/extract-data "resources/smaller Org Chart Data Analysis.xlsx") :version :v1))
  ;; Clara-EAV expects raw EAV records, not transformed vectors

  ;; Session is now defined above, outside the comment block  ;; INCORRECT: Explicit rule vectors don't work properly in Clara-EAV
  (rules/defsession test-session 'clara-org-chart.rules)


  ;; For better performance with large files, use :minimal true
  (def results-fast (-> test-session
                        (rules/insert-all (pos/extract-positions (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx" :minimal true)))
                        (rules/fire-rules)))

  (def results (-> test-session
                   (rules/insert-all (pos/extract-positions (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx")))
                   (rules/fire-rules)))

  (tap> (rules/query results get-position-values))

  (tap> (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx" :minimal true))

  (tap> (pos/extract-positions (xlsx/extract-data "resources/smaller Org Chart Data Analysis.xlsx")))

  (tap> results)
  (tap> (xlsx/extract-data "resources/smaller Org Chart Data Analysis.xlsx"))
  :rcf)