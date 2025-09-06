(ns user
  "Development namespace that's automatically loaded when starting a REPL"
  (:require [portal.api :as p]))

;; Portal setup for development
(defn start-portal! 
  "Start Portal and return the portal instance"
  []
  (def portal (p/open {:theme :portal.colors/nord}))
  (add-tap #'p/submit)
  (println "Portal started! Access at" (str "http://localhost:" (:port portal)))
  portal)

(defn stop-portal! 
  "Stop Portal and remove tap"
  []
  (when (bound? #'portal)
    (remove-tap #'p/submit)
    (p/close portal)
    (println "Portal stopped")))

;; Auto-start Portal when this namespace loads (only in dev)
(when (System/getProperty "user.dir")
  (println "Loading development environment...")
  (start-portal!))

(println "Development utilities loaded:")
(println "  (start-portal!) - Start Portal")
(println "  (stop-portal!)  - Stop Portal") 
(println "  (tap> data)     - Send data to Portal")
