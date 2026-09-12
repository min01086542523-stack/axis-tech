;; mbed.lsp - ARES Commander / AutoCAD
;; (load "C:/Users/ss/manufacturing-mes/cad/mbed.lsp")
;; MBED

(defun mbed-xy (ip x y)
  (list (+ (car ip) x) (+ (cadr ip) y) 0.0)
)

(defun mbed-line (a b)
  (entmake (list (cons 0 "LINE") (cons 8 "0") (cons 10 a) (cons 11 b)))
)

(defun mbed-circ (c r)
  (entmake (list (cons 0 "CIRCLE") (cons 8 "0") (cons 10 c) (cons 40 r)))
)

(defun mbed-draw (pt1 / y33 y313 y563 x42 x57 x122 x582 x647 x662)
  (setq y33 33.0 y313 313.0 y563 563.0
        x42 42.5 x57 57.5 x122 122.5
        x582 582.5 x647 647.5 x662 662.5
  )
  (mbed-line (mbed-xy pt1 0.0 0.0) (mbed-xy pt1 705.0 0.0))
  (mbed-line (mbed-xy pt1 705.0 0.0) (mbed-xy pt1 705.0 596.0))
  (mbed-line (mbed-xy pt1 705.0 596.0) (mbed-xy pt1 0.0 596.0))
  (mbed-line (mbed-xy pt1 0.0 596.0) (mbed-xy pt1 0.0 0.0))
  (mbed-line (mbed-xy pt1 0.0 y33) (mbed-xy pt1 705.0 y33))
  (mbed-line (mbed-xy pt1 0.0 y563) (mbed-xy pt1 705.0 y563))
  (mbed-line (mbed-xy pt1 x57 y33) (mbed-xy pt1 x57 y563))
  (mbed-line (mbed-xy pt1 x647 y33) (mbed-xy pt1 x647 y563))
  (mbed-line (mbed-xy pt1 x122 y33) (mbed-xy pt1 x122 y313))
  (mbed-line (mbed-xy pt1 x582 y33) (mbed-xy pt1 x582 y313))
  (mbed-line (mbed-xy pt1 x57 y313) (mbed-xy pt1 x122 y313))
  (mbed-line (mbed-xy pt1 x582 y313) (mbed-xy pt1 x647 y313))
  (mbed-line (mbed-xy pt1 x122 58.0) (mbed-xy pt1 x582 58.0))
  (mbed-line (mbed-xy pt1 67.5 y313) (mbed-xy pt1 67.5 y563))
  (mbed-line (mbed-xy pt1 637.5 y313) (mbed-xy pt1 637.5 y563))
  (mbed-line (mbed-xy pt1 137.5 y33) (mbed-xy pt1 567.5 y33))
  (mbed-line (mbed-xy pt1 567.5 y33) (mbed-xy pt1 x582 48.0))
  (mbed-line (mbed-xy pt1 x582 48.0) (mbed-xy pt1 x582 y313))
  (mbed-line (mbed-xy pt1 x647 y313) (mbed-xy pt1 x647 548.0))
  (mbed-line (mbed-xy pt1 x647 548.0) (mbed-xy pt1 632.5 y563))
  (mbed-line (mbed-xy pt1 632.5 y563) (mbed-xy pt1 72.5 y563))
  (mbed-line (mbed-xy pt1 72.5 y563) (mbed-xy pt1 x57 548.0))
  (mbed-line (mbed-xy pt1 x57 548.0) (mbed-xy pt1 x57 y313))
  (mbed-line (mbed-xy pt1 x57 y313) (mbed-xy pt1 x122 y313))
  (mbed-line (mbed-xy pt1 x122 y313) (mbed-xy pt1 x122 48.0))
  (mbed-line (mbed-xy pt1 x122 48.0) (mbed-xy pt1 137.5 y33))
  (mbed-circ (mbed-xy pt1 x42 y33) 12.0)
  (mbed-circ (mbed-xy pt1 x662 y33) 12.0)
  (mbed-circ (mbed-xy pt1 x42 y563) 12.0)
  (mbed-circ (mbed-xy pt1 x662 y563) 12.0)
  (mbed-line (mbed-xy pt1 x647 283.0) (mbed-xy pt1 713.0 283.0))
  (mbed-line (mbed-xy pt1 x647 463.0) (mbed-xy pt1 713.0 463.0))
)

(defun c:MBED ( / pt1 cmdecho osmode )
  (setq cmdecho (getvar "CMDECHO"))
  (setq osmode (getvar "OSMODE"))
  (setvar "CMDECHO" 0)
  (setvar "OSMODE" 0)

  (setq pt1 (getpoint "\nClick insert point: "))

  (if pt1
    (progn
      (mbed-draw pt1)
      (princ "\n[MBED] done.")
    )
  )

  (setvar "OSMODE" osmode)
  (setvar "CMDECHO" cmdecho)
  (princ)
)

(princ "\nMBED loaded. Type MBED")
(princ)
