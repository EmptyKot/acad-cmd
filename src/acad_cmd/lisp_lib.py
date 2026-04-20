MCP_DICT_LISP_LIB = r"""(progn
  (defun mcp--json-escape (s / i c out)
        (setq out "")
        (setq i 1)
        (while (<= i (strlen s))
          (setq c (substr s i 1))
          (cond
            ((= c "\\") (setq out (strcat out "\\\\")))
            ((= c "\"") (setq out (strcat out "\\\"")))
            (T (setq out (strcat out c)))
          )
          (setq i (+ i 1))
        )
        out
      )

      (defun mcp--json-quote (s)
        (strcat "\"" (mcp--json-escape s) "\"")
      )

      (defun mcp--json-real (r / s)
        ;; Ensure 0.0 serializes as "0" (not empty string).
        (setq s (vl-string-right-trim "." (vl-string-right-trim "0" (rtos r 2 15))))
        (if (= s "") "0" s)
      )

      (defun mcp--json-value (v)
        (cond
          ((= v T) "true")
          ((= v nil) "false")
          ((and (= (type v) 'SYM) (= (strcase (vl-symbol-name v)) "MCPNULL")) "null")
          ((= (type v) 'STR) (mcp--json-quote v))
          ((= (type v) 'INT) (itoa v))
          ((= (type v) 'REAL) (mcp--json-real v))
          ((= (type v) 'LIST) (mcp--json-arr v))
          (T (mcp--json-quote (vl-princ-to-string v)))
        )
      )

      (defun mcp--json-arr (lst / out first)
        (setq out "[")
        (setq first T)
        (foreach v lst
          (if first
            (setq first nil)
            (setq out (strcat out ","))
          )
          (setq out (strcat out (mcp--json-value v)))
        )
        (setq out (strcat out "]"))
        out
      )

      (defun mcp--emit-json (json)
        (prompt (strcat "\n" "[MCP:JSON]" json))
        (princ)
      )

      (defun mcp--emit-ok (body)
        (mcp--emit-json (strcat "{\"ok\":true" body "}"))
      )

      (defun mcp--emit-err (msg)
        (mcp--emit-json (strcat "{\"ok\":false,\"error\":" (mcp--json-value msg) "}"))
      )

      (defun mcp--nod () (namedobjdict))

      (defun mcp--is-system-name (name / u)
        (setq u (strcase name))
        (or (wcmatch u "ACAD_*") (wcmatch u "AEC_*") (wcmatch u "ADSK_*") (wcmatch u "A$*"))
      )

      (defun mcp--dict-by-name (name / nod r)
        (setq nod (mcp--nod))
        (setq r (dictsearch nod name))
        (if (and r (= (cdr (assoc 0 r)) "DICTIONARY"))
          (cdr (assoc -1 r))
          nil
        )
      )

      (defun mcp--dict-entry-pairs (d / el out key)
        ;; Returns list of (key . ename) from DICTIONARY entity list.
        (setq el (entget d))
        (setq out nil)
        (while el
          (if (= (caar el) 3)
            (progn
              (setq key (cdar el))
              (setq el (cdr el))
              (while (and el (/= (caar el) 350))
                (setq el (cdr el))
              )
              (if el
                (progn
                  (setq out (cons (cons key (cdar el)) out))
                  (setq el (cdr el))
                )
              )
            )
            (setq el (cdr el))
          )
        )
        (reverse out)
      )

      (defun mcp--ensure-dict (name / d)
        (setq d (mcp--dict-by-name name))
        (if d
          d
          (progn
            (setq d (entmakex (list (cons 0 "DICTIONARY") (cons 100 "AcDbDictionary"))))
            (dictadd (mcp--nod) name d)
            d
          )
        )
      )

      (defun mcp--xrec-by-key (d key / r e)
        (setq r (dictsearch d key))
        (if (and r (= (cdr (assoc 0 r)) "XRECORD"))
          (cdr (assoc -1 r))
          nil
        )
      )

      (defun mcp--xrec-filter-pairs (pairs / out)
        (setq out nil)
        (foreach p pairs
          (if (and (numberp (car p))
                   (>= (car p) 1)
                   (/= (car p) 5)
                   (/= (car p) 100)
                   (/= (car p) 102)
                   (/= (car p) 280)
                   (/= (car p) 330)
                   (/= (car p) 360))
            (setq out (cons p out))
          )
        )
        (reverse out)
      )

      (defun mcp--xrec-read (e / pairs)
        (setq pairs (entget e))
        (mcp--xrec-filter-pairs pairs)
      )

      (defun mcp--json-xrec-values (pairs / out first)
        ;; pairs: list of (code . value) -> JSON [[code,value],...]
        (setq out "[")
        (setq first T)
        (foreach p pairs
          (if first
            (setq first nil)
            (setq out (strcat out ","))
          )
          (setq out (strcat out "[" (itoa (car p)) "," (mcp--json-value (cdr p)) "]"))
        )
        (setq out (strcat out "]"))
        out
      )

      (defun mcp--dicts-json (/ nod it out first name obj etype isSys reason)
        (setq nod (mcp--nod))
        (setq it (mcp--dict-entry-pairs nod))
        (setq out "[")
        (setq first T)
        (foreach kv it
          (setq name (car kv))
          (setq obj (cdr kv))
          (setq etype (if obj (cdr (assoc 0 (entget obj))) ""))
          (if (= etype "DICTIONARY")
            (progn
              (setq isSys (mcp--is-system-name name))
              (setq reason (if isSys "prefix" 'MCPNULL))
              (if first
                (setq first nil)
                (setq out (strcat out ","))
              )
              (setq out
                (strcat out
                  "{\"name\":" (mcp--json-value name)
                  ",\"is_system_guess\":" (mcp--json-value isSys)
                  ",\"system_reason\":" (mcp--json-value reason)
                  "}"
                )
              )
            )
          )
        )
        (setq out (strcat out "]"))
        out
      )

      (defun mcp-dict-list ()
        (mcp--emit-json (strcat "{\"ok\":true,\"dicts\":" (mcp--dicts-json) "}"))
      )

      (defun mcp-dict-keys (dictName / d it entries keys first k obj etype)
        (setq d (mcp--dict-by-name dictName))
        (if (not d)
          (mcp--emit-json "{\"ok\":true,\"found\":false,\"keys\":[],\"entries\":[]}")
          (progn
            (setq it (mcp--dict-entry-pairs d))
            (setq entries "[")
            (setq keys "[")
            (setq first T)
            (foreach kv it
              (setq k (car kv))
              (setq obj (cdr kv))
              (setq etype (if obj (cdr (assoc 0 (entget obj))) 'MCPNULL))
              (if first
                (setq first nil)
                (progn
                  (setq entries (strcat entries ","))
                  (setq keys (strcat keys ","))
                )
              )
              (setq entries (strcat entries "{\"key\":" (mcp--json-value k) ",\"type\":" (mcp--json-value etype) "}"))
              (setq keys (strcat keys (mcp--json-value k)))
            )
            (setq entries (strcat entries "]"))
            (setq keys (strcat keys "]"))
            (mcp--emit-json (strcat "{\"ok\":true,\"found\":true,\"keys\":" keys ",\"entries\":" entries "}"))
          )
        )
      )

      (defun mcp-xrecord-get (dictName key / d x pairs)
        (setq d (mcp--dict-by-name dictName))
        (if (not d)
          (mcp--emit-json "{\"ok\":true,\"found\":false,\"values\":[]}")
          (progn
            (setq x (mcp--xrec-by-key d key))
            (if (not x)
              (mcp--emit-json "{\"ok\":true,\"found\":false,\"values\":[]}")
              (progn
                (setq pairs (mcp--xrec-read x))
                (mcp--emit-json (strcat "{\"ok\":true,\"found\":true,\"values\":" (mcp--json-xrec-values pairs) "}"))
              )
            )
          )
        )
      )

      (defun mcp-xrecord-set (dictName key values overwrite / d old xrec)
        (setq d (mcp--ensure-dict dictName))
        (setq old (mcp--xrec-by-key d key))
        (if old
          (if overwrite
            (progn
              (dictremove d key)
              (entdel old)
            )
            (progn
              (mcp--emit-err "Key already exists")
              (setq d nil)
            )
          )
        )
        (if d
          (progn
            (setq xrec (entmakex (append (list (cons 0 "XRECORD") (cons 100 "AcDbXrecord")) values)))
            (dictadd d key xrec)
            (mcp--emit-json "{\"ok\":true,\"written\":true}")
          )
        )
      )

      (defun mcp-xrecord-delete (dictName key / d old)
        (setq d (mcp--dict-by-name dictName))
        (if (not d)
          (mcp--emit-json "{\"ok\":true,\"deleted\":false}")
          (progn
            (setq old (mcp--xrec-by-key d key))
            (if (not old)
              (mcp--emit-json "{\"ok\":true,\"deleted\":false}")
              (progn
                (dictremove d key)
                (entdel old)
                (mcp--emit-json "{\"ok\":true,\"deleted\":true}")
              )
            )
          )
        )
      )

  (defun mcp-dict-delete (dictName recursive / nod d it k obj n)
        (setq nod (mcp--nod))
        (setq d (mcp--dict-by-name dictName))
        (if (not d)
          (mcp--emit-json "{\"ok\":true,\"deleted\":false,\"deleted_entries\":0}")
          (progn
            (setq n 0)
            (setq it (mcp--dict-entry-pairs d))
            (if (and (not recursive) it)
              (progn
                (mcp--emit-err "Dictionary not empty (set recursive=true to delete)")
                (setq d nil)
              )
            )
            (if d
              (progn
                (foreach kv it
                  (setq k (car kv))
                  (setq obj (cdr kv))
                  (if k (dictremove d k))
                  (if obj (entdel obj))
                  (setq n (+ n 1))
                )
                (dictremove nod dictName)
                (entdel d)
                (mcp--emit-json (strcat "{\"ok\":true,\"deleted\":true,\"deleted_entries\":" (itoa n) "}"))
              )
            )
          )
        )
  )
  (princ)
 )
"""


MCP_SELECTION_LISP_LIB = MCP_DICT_LISP_LIB + r"""(progn
  (vl-load-com)

      (defun mcp--emit-sel-start (req_id count errno)
        (mcp--emit-json
          (strcat
            "{\"ok\":true,\"req_id\":" (mcp--json-value req_id)
            ",\"event\":\"start\""
            ",\"count\":" (itoa count)
            ",\"errno\":" (itoa errno)
            "}"
          )
        )
      )

      (defun mcp--emit-sel-item-begin-lite (req_id i handle etype)
        (mcp--emit-json
          (strcat
            "{\"ok\":true,\"req_id\":" (mcp--json-value req_id)
            ",\"event\":\"item_begin\""
            ",\"i\":" (itoa i)
            ",\"handle\":" (mcp--json-value handle)
            ",\"type\":" (mcp--json-value etype)
            "}"
          )
        )
      )

      (defun mcp--emit-sel-done (req_id)
        (mcp--emit-json
          (strcat
            "{\"ok\":true,\"req_id\":" (mcp--json-value req_id)
            ",\"event\":\"done\"}"
          )
        )
      )

      (defun mcp-selection--emit-from-ss-lite (req_id ss max_objects / errno total n i ename el handle etype)
        (setq errno (getvar "ERRNO"))
        (setq total (if ss (sslength ss) 0))
        (setq n total)
        (if (and max_objects (> max_objects 0) (> n max_objects))
          (setq n max_objects)
        )
        (mcp--emit-sel-start req_id n errno)
        (setq i 0)
        (while (< i n)
          (setq ename (ssname ss i))
          (setq el (entget ename))
          (setq handle (cdr (assoc 5 el)))
          (setq etype (cdr (assoc 0 el)))
          (mcp--emit-sel-item-begin-lite req_id i handle etype)
          (setq i (+ i 1))
        )
        (mcp--emit-sel-done req_id)
      )

      (defun mcp-selection-implied-lite (req_id max_objects / ss)
        ;; Implied (PickFirst) selection only; never prompt the user.
        (setq ss (ssget "_I"))
        (mcp-selection--emit-from-ss-lite req_id ss max_objects)
      )

      (defun mcp-selection-prompt-lite (req_id prompt_str filter_list max_objects / ss)
        ;; Interactive selection set (user picks in UI).
        (if prompt_str (prompt (strcat "\n" prompt_str)))
        (if filter_list
          (setq ss (ssget filter_list))
          (setq ss (ssget))
        )
        (mcp-selection--emit-from-ss-lite req_id ss max_objects)
      )
  (princ)
)
"""


def lisp_string(s: str) -> str:
    from .lisp import lisp_quote_string

    return '"' + lisp_quote_string(s) + '"'


def lisp_concat(prefix: str, suffix: str) -> str:
    """Concatenate LISP snippets with exactly one newline between."""

    return prefix.rstrip("\r\n") + "\n" + suffix.lstrip("\r\n")


def lisp_typed_values(values):
    """Convert [{code,value},...] into a LISP list of dotted pairs."""

    if values is None:
        return "'()"
    if not isinstance(values, list):
        raise ValueError("values must be a list")

    parts = []
    for i, item in enumerate(values):
        if not isinstance(item, dict):
            raise ValueError(f"values[{i}] must be an object")
        if "code" not in item or "value" not in item:
            raise ValueError(f"values[{i}] must have 'code' and 'value'")
        code = item["code"]
        val = item["value"]
        if not isinstance(code, int):
            raise ValueError(f"values[{i}].code must be integer")

        if isinstance(val, str):
            v = lisp_string(val)
        elif isinstance(val, bool):
            v = "T" if val else "nil"
        elif isinstance(val, int) or isinstance(val, float):
            v = str(val)
        elif isinstance(val, (list, tuple)):
            nums = []
            for j, n in enumerate(val):
                if not isinstance(n, (int, float)):
                    raise ValueError(f"values[{i}].value[{j}] must be number")
                nums.append(str(float(n)))
            v = "(" + " ".join(nums) + ")"
        elif val is None:
            v = "nil"
        else:
            raise ValueError(f"values[{i}].value has unsupported type")

        parts.append(f"(cons {code} {v})")

    return "(list " + " ".join(parts) + ")"


def strip_ok(obj):
    if "ok" not in obj:
        return obj
    out = dict(obj)
    out.pop("ok", None)
    return out
