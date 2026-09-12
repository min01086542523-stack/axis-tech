;; mbed.lsp
;; (load "C:/Users/ss/manufacturing-mes/cad/mbed.lsp")
;; Type: MBED
;; One click -> draw once -> exit. No while. No command. No IMAGE.

(setq *MBED-BUSY* nil)

(defun mbed-xy (ip x y)
  (list (+ (car ip) x) (+ (cadr ip) y) 0.0)
)

(defun mbed-line (a b)
  (entmake (list (cons 0 "LINE") (cons 8 "0") (cons 10 a) (cons 11 b)))
)

(defun mbed-circ (c r)
  (entmake (list (cons 0 "CIRCLE") (cons 8 "0") (cons 10 c) (cons 40 r)))
)

(defun mbed-draw (p)
  (mbed-line (mbed-xy p 0.0 0.0) (mbed-xy p 705.0 0.0))
  (mbed-line (mbed-xy p 705.0 0.0) (mbed-xy p 705.0 596.0))
  (mbed-line (mbed-xy p 705.0 596.0) (mbed-xy p 0.0 596.0))
  (mbed-line (mbed-xy p 0.0 596.0) (mbed-xy p 0.0 0.0))
  (mbed-line (mbed-xy p 0.0 33.0) (mbed-xy p 705.0 33.0))
  (mbed-line (mbed-xy p 0.0 563.0) (mbed-xy p 705.0 563.0))
  (mbed-line (mbed-xy p 57.5 33.0) (mbed-xy p 57.5 563.0))
  (mbed-line (mbed-xy p 647.5 33.0) (mbed-xy p 647.5 563.0))
  (mbed-line (mbed-xy p 122.5 33.0) (mbed-xy p 122.5 313.0))
  (mbed-line (mbed-xy p 582.5 33.0) (mbed-xy p 582.5 313.0))
  (mbed-line (mbed-xy p 57.5 313.0) (mbed-xy p 122.5 313.0))
  (mbed-line (mbed-xy p 582.5 313.0) (mbed-xy p 647.5 313.0))
  (mbed-circ (mbed-xy p 42.5 33.0) 12.0)
  (mbed-circ (mbed-xy p 662.5 33.0) 12.0)
  (mbed-circ (mbed-xy p 42.5 563.0) 12.0)
  (mbed-circ (mbed-xy p 662.5 563.0) 12.0)
)

(defun c:MBED ( / pt1 echo os )
  (if *MBED-BUSY*
    (princ "\nMBED busy.")
    (progn
      (setq *MBED-BUSY* T)
      (setq echo (getvar "CMDECHO"))
      (setq os (getvar "OSMODE"))
      (setvar "CMDECHO" 0)
      (setvar "OSMODE" 0)
      (setq pt1 (getpoint "\nClick once: "))
      (if pt1
        (mbed-draw pt1)
      )
      (setvar "OSMODE" os)
      (setvar "CMDECHO" echo)
      (setq *MBED-BUSY* nil)
      (princ "\n[MBED] done.")
    )
  )
  (princ)
)

(princ "\nType MBED then click once.")
(princ)
