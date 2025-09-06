(ns clara-org-chart.position
  (:require [clara.rules :refer :all])
  )


;; Position	Title	Current Employee	Position Reports To Position	Position Reports To Position (Dotted Line)	City	Comments/Notes	Dotted_Reports	AgencyCode	UnitCode	ClassCode	SerialNumber	Unique_Flag	Time_Base	TB_Adjustment	Region
(defrecord Position 
           [
            position
            title
            current-employee
            reports-to-position
            dotted-line-reports-to-position
            city
            comments-notes
            dotted-reports
            agencycode
            unitcode
            classcode
            serialnumber
            unique-flag
            time-base
            tb-adjustment
            region
           ])




(defn extract-positions [xlsx-data]
  (let [sheet (first (filter #(= (:sheet-name %) "ALL REGIONS") (:sheets xlsx-data)))
    cells (:cells sheet)
    headers (->> cells
         (filter #(= (:row %) 0))
         (sort-by :col)
         (map :value))
    rows (->> cells
      (remove #(= (:row %) 0))
      (group-by :row))]
    (map (fn [[_ row-cells]]
       (let [row-map (->> row-cells
          (sort-by :col)
          (map :value)
          (zipmap headers))]
     (map->Position
      {:position (get row-map "Position")
       :title (get row-map "Title")
       :current-employee (get row-map "Current Employee")
       :reports-to-position (get row-map "Position Reports To Position")
       :dotted-line-reports-to-position (get row-map "Position Reports To Position (Dotted Line)")
       :city (get row-map "City")
       :comments-notes (get row-map "Comments/Notes")
       :dotted-reports (get row-map "Dotted_Reports")
       :agencycode (get row-map "AgencyCode")
       :unitcode (get row-map "UnitCode")
       :classcode (get row-map "ClassCode")
       :serialnumber (get row-map "SerialNumber")
       :unique-flag (get row-map "Unique_Flag")
       :time-base (get row-map "Time_Base")
       :tb-adjustment (get row-map "TB_Adjustment")
       :region (get row-map "Region")})))
     rows)))