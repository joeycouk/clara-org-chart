(ns data-types)

(defrecord Position
           [row-num
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
            tenure
            region
            position-weight
            total-subordinates  ; New field for subordinate count
            ])