#pragma once

#include once "crt/stdint.bi"
#include once "crt/stdio.bi"
#include once "crt/errno.bi"
#include once "crt/stdlib.bi"
#include once "crt/string.bi"
#include once "crt/time.bi"

extern "C"
#define NULL 0

#define __LXW_XLSXWRITER_H__
#define __LXW_WORKBOOK_H__
#define __LXW_WORKSHEET_H__
#define __LXW_SST_H__
#define __LXW_COMMON_H__
#define _SYS_QUEUE_H_
#define QMD_TRACE_ELEM(elem)
#define QMD_TRACE_HEAD(head)
#define QMD_SAVELINK(name, link)
#define TRACEBUF
#define TRACEBUF_INITIALIZER
#define TRASHIT(x)
'' TODO: #define SLIST_HEAD(name, type) struct name { struct type *slh_first; }
#define SLIST_HEAD_INITIALIZER(head) (NULL)
'' TODO: #define SLIST_ENTRY(type) struct { struct type *sle_next; }
#define SLIST_EMPTY(head) ((head)->slh_first = NULL)
#define SLIST_FIRST(head) (head)->slh_first

'' TODO: #define SLIST_FOREACH(var, head, field) for ((var) = SLIST_FIRST((head)); (var); (var) = SLIST_NEXT((var), field))
'' TODO: #define SLIST_FOREACH_FROM(var, head, field) for ((var) = ((var) ? (var) : SLIST_FIRST((head))); (var); (var) = SLIST_NEXT((var), field))
'' TODO: #define SLIST_FOREACH_SAFE(var, head, field, tvar) for ((var) = SLIST_FIRST((head)); (var) && ((tvar) = SLIST_NEXT((var), field), 1); (var) = (tvar))
'' TODO: #define SLIST_FOREACH_FROM_SAFE(var, head, field, tvar) for ((var) = ((var) ? (var) : SLIST_FIRST((head))); (var) && ((tvar) = SLIST_NEXT((var), field), 1); (var) = (tvar))
'' TODO: #define SLIST_FOREACH_PREVPTR(var, varp, head, field) for ((varp) = &SLIST_FIRST((head)); ((var) = *(varp)) != NULL; (varp) = &SLIST_NEXT((var), field))

#define SLIST_INIT(head) scope : SLIST_FIRST((head)) = NULL : end scope
#macro SLIST_INSERT_AFTER(slistelm, elm, field)
	scope
		SLIST_NEXT((elm), field) = SLIST_NEXT((slistelm), field)
		SLIST_NEXT((slistelm), field) = (elm)
	end scope
#endmacro
#macro SLIST_INSERT_HEAD(head, elm, field)
	scope
		SLIST_NEXT((elm), field) = SLIST_FIRST((head))
		SLIST_FIRST((head)) = (elm)
	end scope
#endmacro
#define SLIST_NEXT(elm, field) (elm)->field.sle_next
#macro SLIST_REMOVE(head, elm, type, field)
	scope
		QMD_SAVELINK(oldnext, (elm)->field.sle_next)
		if SLIST_FIRST((head)) = (elm) then
			SLIST_REMOVE_HEAD((head), field)
		else
			dim curelm as type ptr = SLIST_FIRST((head))
			while SLIST_NEXT(curelm, field) <> (elm)
				curelm = SLIST_NEXT(curelm, field)
			wend
			SLIST_REMOVE_AFTER(curelm, field)
		end if
		TRASHIT(*oldnext)
	end scope
#endmacro
#define SLIST_REMOVE_AFTER(elm, field) scope : SLIST_NEXT(elm, field) = SLIST_NEXT(SLIST_NEXT(elm, field), field) : end scope
#define SLIST_REMOVE_HEAD(head, field) scope : SLIST_FIRST((head)) = SLIST_NEXT(SLIST_FIRST((head)), field) : end scope
#macro SLIST_SWAP(head1, head2, type)
	scope
		dim swap_first as type ptr = SLIST_FIRST(head1)
		SLIST_FIRST(head1) = SLIST_FIRST(head2)
		SLIST_FIRST(head2) = swap_first
	end scope
#endmacro
#define STAILQ_HEAD(name, type) struct name { struct type *stqh_first; struct type **stqh_last; }
#define STAILQ_HEAD_INITIALIZER(head) (NULL, @(head).stqh_first)
#define STAILQ_ENTRY(type) struct { struct type *stqe_next; }
#macro STAILQ_CONCAT(head1, head2)
	if STAILQ_EMPTY((head2)) = 0 then
		(*(head1)->stqh_last) = (head2)->stqh_first
		(head1)->stqh_last = (head2)->stqh_last
		STAILQ_INIT((head2))
	end if
#endmacro
#define STAILQ_EMPTY(head) ((head)->stqh_first = NULL)
#define STAILQ_FIRST(head) (head)->stqh_first

#define STAILQ_FOREACH(var, head, field) for((var) = STAILQ_FIRST((head)); (var); (var) = STAILQ_NEXT((var), field))
#define STAILQ_FOREACH_FROM(var, head, field) for ((var) = ((var) ? (var) : STAILQ_FIRST((head))); (var); (var) = STAILQ_NEXT((var), field))
#define STAILQ_FOREACH_SAFE(var, head, field, tvar) for ((var) = STAILQ_FIRST((head)); (var) && ((tvar) = STAILQ_NEXT((var), field), 1); (var) = (tvar))
#define STAILQ_FOREACH_FROM_SAFE(var, head, field, tvar) for ((var) = ((var) ? (var) : STAILQ_FIRST((head))); (var) && ((tvar) = STAILQ_NEXT((var), field), 1); (var) = (tvar))

#macro STAILQ_INIT(head)
	scope
		STAILQ_FIRST((head)) = NULL
		(head)->stqh_last = @STAILQ_FIRST((head))
	end scope
#endmacro
#macro STAILQ_INSERT_AFTER(head, tqelm, elm, field)
	scope
		'' TODO: if ((STAILQ_NEXT((elm), field) = STAILQ_NEXT((tqelm), field)) == NULL) (head)->stqh_last = &STAILQ_NEXT((elm), field);
		STAILQ_NEXT((tqelm), field) = (elm)
	end scope
#endmacro
#macro STAILQ_INSERT_HEAD(head, elm, field)
	scope
		'' TODO: if ((STAILQ_NEXT((elm), field) = STAILQ_FIRST((head))) == NULL) (head)->stqh_last = &STAILQ_NEXT((elm), field);
		STAILQ_FIRST((head)) = (elm)
	end scope
#endmacro
#macro STAILQ_INSERT_TAIL(head, elm, field)
	scope
		STAILQ_NEXT((elm), field) = NULL
		(*(head)->stqh_last) = (elm)
		(head)->stqh_last = @STAILQ_NEXT((elm), field)
	end scope
#endmacro
'' TODO: #define STAILQ_LAST(head, type, field) (STAILQ_EMPTY((head)) ? NULL : __containerof((head)->stqh_last, struct type, field.stqe_next))
#define STAILQ_NEXT(elm, field) (elm)->field.stqe_next
#macro STAILQ_REMOVE(head, elm, type, field)
	scope
		QMD_SAVELINK(oldnext, (elm)->field.stqe_next)
		if STAILQ_FIRST((head)) = (elm) then
			STAILQ_REMOVE_HEAD((head), field)
		else
			dim curelm as type ptr = STAILQ_FIRST((head))
			while STAILQ_NEXT(curelm, field) <> (elm)
				curelm = STAILQ_NEXT(curelm, field)
			wend
			STAILQ_REMOVE_AFTER(head, curelm, field)
		end if
		TRASHIT(*oldnext)
	end scope
#endmacro
#define STAILQ_REMOVE_AFTER(head, elm, field) scope : /' TODO: if ((STAILQ_NEXT(elm, field) = STAILQ_NEXT(STAILQ_NEXT(elm, field), field)) == NULL) (head)->stqh_last = &STAILQ_NEXT((elm), field); '/ : end scope
#define STAILQ_REMOVE_HEAD(head, field) scope : /' TODO: if ((STAILQ_FIRST((head)) = STAILQ_NEXT(STAILQ_FIRST((head)), field)) == NULL) (head)->stqh_last = &STAILQ_FIRST((head)); '/ : end scope
#macro STAILQ_SWAP(head1, head2, type)
	scope
		dim swap_first as type ptr = STAILQ_FIRST(head1)
		dim swap_last as type ptr ptr = (head1)->stqh_last
		STAILQ_FIRST(head1) = STAILQ_FIRST(head2)
		(head1)->stqh_last = (head2)->stqh_last
		STAILQ_FIRST(head2) = swap_first
		(head2)->stqh_last = swap_last
		if STAILQ_EMPTY(head1) then
			(head1)->stqh_last = @STAILQ_FIRST(head1)
		end if
		if STAILQ_EMPTY(head2) then
			(head2)->stqh_last = @STAILQ_FIRST(head2)
		end if
	end scope
#endmacro
'' TODO: #define LIST_HEAD(name, type) struct name { struct type *lh_first; }
#define LIST_HEAD_INITIALIZER(head) (NULL)
'' TODO: #define LIST_ENTRY(type) struct { struct type *le_next; struct type **le_prev; }
#define QMD_LIST_CHECK_HEAD(head, field)
#define QMD_LIST_CHECK_NEXT(elm, field)
#define QMD_LIST_CHECK_PREV(elm, field)
#define LIST_EMPTY(head) ((head)->lh_first = NULL)
#define LIST_FIRST(head) (head)->lh_first

'' TODO: #define LIST_FOREACH(var, head, field) for ((var) = LIST_FIRST((head)); (var); (var) = LIST_NEXT((var), field))
'' TODO: #define LIST_FOREACH_FROM(var, head, field) for ((var) = ((var) ? (var) : LIST_FIRST((head))); (var); (var) = LIST_NEXT((var), field))
'' TODO: #define LIST_FOREACH_SAFE(var, head, field, tvar) for ((var) = LIST_FIRST((head)); (var) && ((tvar) = LIST_NEXT((var), field), 1); (var) = (tvar))
'' TODO: #define LIST_FOREACH_FROM_SAFE(var, head, field, tvar) for ((var) = ((var) ? (var) : LIST_FIRST((head))); (var) && ((tvar) = LIST_NEXT((var), field), 1); (var) = (tvar))

#define LIST_INIT(head) scope : LIST_FIRST((head)) = NULL : end scope
#macro LIST_INSERT_AFTER(listelm, elm, field)
	scope
		QMD_LIST_CHECK_NEXT(listelm, field)
		'' TODO: if ((LIST_NEXT((elm), field) = LIST_NEXT((listelm), field)) != NULL) LIST_NEXT((listelm), field)->field.le_prev = &LIST_NEXT((elm), field);
		LIST_NEXT((listelm), field) = (elm)
		(elm)->field.le_prev = @LIST_NEXT((listelm), field)
	end scope
#endmacro
#macro LIST_INSERT_BEFORE(listelm, elm, field)
	scope
		QMD_LIST_CHECK_PREV(listelm, field)
		(elm)->field.le_prev = (listelm)->field.le_prev
		LIST_NEXT((elm), field) = (listelm)
		(*(listelm)->field.le_prev) = (elm)
		(listelm)->field.le_prev = @LIST_NEXT((elm), field)
	end scope
#endmacro
#macro LIST_INSERT_HEAD(head, elm, field)
	scope
		QMD_LIST_CHECK_HEAD((head), field)
		'' TODO: if ((LIST_NEXT((elm), field) = LIST_FIRST((head))) != NULL) LIST_FIRST((head))->field.le_prev = &LIST_NEXT((elm), field);
		LIST_FIRST((head)) = (elm)
		(elm)->field.le_prev = @LIST_FIRST((head))
	end scope
#endmacro
#define LIST_NEXT(elm, field) (elm)->field.le_next
'' TODO: #define LIST_PREV(elm, head, type, field) ((elm)->field.le_prev == &LIST_FIRST((head)) ? NULL : __containerof((elm)->field.le_prev, struct type, field.le_next))
#macro LIST_REMOVE(elm, field)
	scope
		QMD_SAVELINK(oldnext, (elm)->field.le_next)
		QMD_SAVELINK(oldprev, (elm)->field.le_prev)
		QMD_LIST_CHECK_NEXT(elm, field)
		QMD_LIST_CHECK_PREV(elm, field)
		if LIST_NEXT((elm), field) <> NULL then
			LIST_NEXT((elm), field)->field.le_prev = (elm)->field.le_prev
		end if
		(*(elm)->field.le_prev) = LIST_NEXT((elm), field)
		TRASHIT(*oldnext)
		TRASHIT(*oldprev)
	end scope
#endmacro
#macro LIST_SWAP(head1, head2, type, field)
	scope
		dim swap_tmp as type ptr = LIST_FIRST((head1))
		LIST_FIRST((head1)) = LIST_FIRST((head2))
		LIST_FIRST((head2)) = swap_tmp
		'' TODO: if ((swap_tmp = LIST_FIRST((head1))) != NULL) swap_tmp->field.le_prev = &LIST_FIRST((head1));
		'' TODO: if ((swap_tmp = LIST_FIRST((head2))) != NULL) swap_tmp->field.le_prev = &LIST_FIRST((head2));
	end scope
#endmacro
'' TODO: #define TAILQ_HEAD(name, type) struct name { struct type *tqh_first; struct type **tqh_last; TRACEBUF }
#define TAILQ_HEAD_INITIALIZER(head) (NULL, @(head).tqh_first, TRACEBUF_INITIALIZER)
'' TODO: #define TAILQ_ENTRY(type) struct { struct type *tqe_next; struct type **tqe_prev; TRACEBUF }
#define QMD_TAILQ_CHECK_HEAD(head, field)
#define QMD_TAILQ_CHECK_TAIL(head, headname)
#define QMD_TAILQ_CHECK_NEXT(elm, field)
#define QMD_TAILQ_CHECK_PREV(elm, field)
#macro TAILQ_CONCAT(head1, head2, field)
	if TAILQ_EMPTY(head2) = 0 then
		(*(head1)->tqh_last) = (head2)->tqh_first
		(head2)->tqh_first->field.tqe_prev = (head1)->tqh_last
		(head1)->tqh_last = (head2)->tqh_last
		TAILQ_INIT((head2))
		QMD_TRACE_HEAD(head1)
		QMD_TRACE_HEAD(head2)
	end if
#endmacro
#define TAILQ_EMPTY(head) ((head)->tqh_first = NULL)
#define TAILQ_FIRST(head) (head)->tqh_first

'' TODO: #define TAILQ_FOREACH(var, head, field) for ((var) = TAILQ_FIRST((head)); (var); (var) = TAILQ_NEXT((var), field))
'' TODO: #define TAILQ_FOREACH_FROM(var, head, field) for ((var) = ((var) ? (var) : TAILQ_FIRST((head))); (var); (var) = TAILQ_NEXT((var), field))
'' TODO: #define TAILQ_FOREACH_SAFE(var, head, field, tvar) for ((var) = TAILQ_FIRST((head)); (var) && ((tvar) = TAILQ_NEXT((var), field), 1); (var) = (tvar))
'' TODO: #define TAILQ_FOREACH_FROM_SAFE(var, head, field, tvar) for ((var) = ((var) ? (var) : TAILQ_FIRST((head))); (var) && ((tvar) = TAILQ_NEXT((var), field), 1); (var) = (tvar))
'' TODO: #define TAILQ_FOREACH_REVERSE(var, head, headname, field) for ((var) = TAILQ_LAST((head), headname); (var); (var) = TAILQ_PREV((var), headname, field))
'' TODO: #define TAILQ_FOREACH_REVERSE_FROM(var, head, headname, field) for ((var) = ((var) ? (var) : TAILQ_LAST((head), headname)); (var); (var) = TAILQ_PREV((var), headname, field))
'' TODO: #define TAILQ_FOREACH_REVERSE_SAFE(var, head, headname, field, tvar) for ((var) = TAILQ_LAST((head), headname); (var) && ((tvar) = TAILQ_PREV((var), headname, field), 1); (var) = (tvar))
'' TODO: #define TAILQ_FOREACH_REVERSE_FROM_SAFE(var, head, headname, field, tvar) for ((var) = ((var) ? (var) : TAILQ_LAST((head), headname)); (var) && ((tvar) = TAILQ_PREV((var), headname, field), 1); (var) = (tvar))

#macro TAILQ_INIT(head)
	scope
		TAILQ_FIRST((head)) = NULL
		(head)->tqh_last = @TAILQ_FIRST((head))
		QMD_TRACE_HEAD(head)
	end scope
#endmacro
#macro TAILQ_INSERT_AFTER(head, listelm, elm, field)
	scope
		QMD_TAILQ_CHECK_NEXT(listelm, field)
		'' TODO: if ((TAILQ_NEXT((elm), field) = TAILQ_NEXT((listelm), field)) != NULL) TAILQ_NEXT((elm), field)->field.tqe_prev = &TAILQ_NEXT((elm), field);
		'' TODO: else { (head)->tqh_last = &TAILQ_NEXT((elm), field); QMD_TRACE_HEAD(head); } TAILQ_NEXT((listelm), field) = (elm);
		(elm)->field.tqe_prev = @TAILQ_NEXT((listelm), field)
		QMD_TRACE_ELEM(@(elm)->field)
		QMD_TRACE_ELEM(@listelm->field)
	end scope
#endmacro
#macro TAILQ_INSERT_BEFORE(listelm, elm, field)
	scope
		QMD_TAILQ_CHECK_PREV(listelm, field)
		(elm)->field.tqe_prev = (listelm)->field.tqe_prev
		TAILQ_NEXT((elm), field) = (listelm)
		(*(listelm)->field.tqe_prev) = (elm)
		(listelm)->field.tqe_prev = @TAILQ_NEXT((elm), field)
		QMD_TRACE_ELEM(@(elm)->field)
		QMD_TRACE_ELEM(@listelm->field)
	end scope
#endmacro
#macro TAILQ_INSERT_HEAD(head, elm, field)
	scope
		QMD_TAILQ_CHECK_HEAD(head, field)
		'' TODO: if ((TAILQ_NEXT((elm), field) = TAILQ_FIRST((head))) != NULL) TAILQ_FIRST((head))->field.tqe_prev = &TAILQ_NEXT((elm), field);
		'' TODO: else (head)->tqh_last = &TAILQ_NEXT((elm), field);
		TAILQ_FIRST((head)) = (elm)
		(elm)->field.tqe_prev = @TAILQ_FIRST((head))
		QMD_TRACE_HEAD(head)
		QMD_TRACE_ELEM(@(elm)->field)
	end scope
#endmacro
#macro TAILQ_INSERT_TAIL(head, elm, field)
	scope
		QMD_TAILQ_CHECK_TAIL(head, field)
		TAILQ_NEXT((elm), field) = NULL
		(elm)->field.tqe_prev = (head)->tqh_last
		(*(head)->tqh_last) = (elm)
		(head)->tqh_last = @TAILQ_NEXT((elm), field)
		QMD_TRACE_HEAD(head)
		QMD_TRACE_ELEM(@(elm)->field)
	end scope
#endmacro
#define TAILQ_LAST(head, headname) (*cptr(headname ptr, (head)->tqh_last)->tqh_last)
#define TAILQ_NEXT(elm, field) (elm)->field.tqe_next
#define TAILQ_PREV(elm, headname, field) (*cptr(headname ptr, (elm)->field.tqe_prev)->tqh_last)
#macro TAILQ_REMOVE(head, elm, field)
	scope
		QMD_SAVELINK(oldnext, (elm)->field.tqe_next)
		QMD_SAVELINK(oldprev, (elm)->field.tqe_prev)
		QMD_TAILQ_CHECK_NEXT(elm, field)
		QMD_TAILQ_CHECK_PREV(elm, field)
		if TAILQ_NEXT((elm), field) <> NULL then
			TAILQ_NEXT((elm), field)->field.tqe_prev = (elm)->field.tqe_prev
		else
			(head)->tqh_last = (elm)->field.tqe_prev
			QMD_TRACE_HEAD(head)
		end if
		(*(elm)->field.tqe_prev) = TAILQ_NEXT((elm), field)
		TRASHIT(*oldnext)
		TRASHIT(*oldprev)
		QMD_TRACE_ELEM(@(elm)->field)
	end scope
#endmacro
#macro TAILQ_SWAP(head1, head2, type, field)
	scope
		dim swap_first as type ptr = (head1)->tqh_first
		dim swap_last as type ptr ptr = (head1)->tqh_last
		(head1)->tqh_first = (head2)->tqh_first
		(head1)->tqh_last = (head2)->tqh_last
		(head2)->tqh_first = swap_first
		(head2)->tqh_last = swap_last
		'' TODO: if ((swap_first = (head1)->tqh_first) != NULL) swap_first->field.tqe_prev = &(head1)->tqh_first;
		'' TODO: else (head1)->tqh_last = &(head1)->tqh_first;
		'' TODO: if ((swap_first = (head2)->tqh_first) != NULL) swap_first->field.tqe_prev = &(head2)->tqh_first;
		'' TODO: else (head2)->tqh_last = &(head2)->tqh_first;
	end scope
#endmacro
#define _SYS_TREE_H_
'' TODO: #define SPLAY_HEAD(name, type) struct name { struct type *sph_root; }
#define SPLAY_INITIALIZER(root) (NULL)
#define SPLAY_INIT(root) scope : (root)->sph_root = NULL : end scope
'' TODO: #define SPLAY_ENTRY(type) struct { struct type *spe_left; struct type *spe_right; }
#define SPLAY_LEFT(elm, field) (elm)->field.spe_left
#define SPLAY_RIGHT(elm, field) (elm)->field.spe_right
#define SPLAY_ROOT(head) (head)->sph_root
#define SPLAY_EMPTY(head) (SPLAY_ROOT(head) = NULL)
#macro SPLAY_ROTATE_RIGHT(head, tmp, field)
	scope
		SPLAY_LEFT((head)->sph_root, field) = SPLAY_RIGHT(tmp, field)
		SPLAY_RIGHT(tmp, field) = (head)->sph_root
		(head)->sph_root = tmp
	end scope
#endmacro
#macro SPLAY_ROTATE_LEFT(head, tmp, field)
	scope
		SPLAY_RIGHT((head)->sph_root, field) = SPLAY_LEFT(tmp, field)
		SPLAY_LEFT(tmp, field) = (head)->sph_root
		(head)->sph_root = tmp
	end scope
#endmacro
#macro SPLAY_LINKLEFT(head, tmp, field)
	scope
		SPLAY_LEFT(tmp, field) = (head)->sph_root
		tmp = (head)->sph_root
		(head)->sph_root = SPLAY_LEFT((head)->sph_root, field)
	end scope
#endmacro
#macro SPLAY_LINKRIGHT(head, tmp, field)
	scope
		SPLAY_RIGHT(tmp, field) = (head)->sph_root
		tmp = (head)->sph_root
		(head)->sph_root = SPLAY_RIGHT((head)->sph_root, field)
	end scope
#endmacro
#macro SPLAY_ASSEMBLE(head, node, left, right, field)
	scope
		SPLAY_RIGHT(left, field) = SPLAY_LEFT((head)->sph_root, field)
		SPLAY_LEFT(right, field) = SPLAY_RIGHT((head)->sph_root, field)
		SPLAY_LEFT((head)->sph_root, field) = SPLAY_RIGHT(node, field)
		SPLAY_RIGHT((head)->sph_root, field) = SPLAY_LEFT(node, field)
	end scope
#endmacro
'' TODO: #define SPLAY_PROTOTYPE(name, type, field, cmp) void name##_SPLAY(struct name *, struct type *); void name##_SPLAY_MINMAX(struct name *, int); struct type *name##_SPLAY_INSERT(struct name *, struct type *); struct type *name##_SPLAY_REMOVE(struct name *, struct type *); static __inline struct type * name##_SPLAY_FIND(struct name *head, struct type *elm) { if (SPLAY_EMPTY(head)) return(NULL); name##_SPLAY(head, elm); if ((cmp)(elm, (head)->sph_root) == 0) return (head->sph_root); return (NULL); } static __inline struct type * name##_SPLAY_NEXT(struct name *head, struct type *elm) { name##_SPLAY(head, elm); if (SPLAY_RIGHT(elm, field) != NULL) { elm = SPLAY_RIGHT(elm, field); while (SPLAY_LEFT(elm, field) != NULL) { elm = SPLAY_LEFT(elm, field); } } else elm = NULL; return (elm); } static __inline struct type * name##_SPLAY_MIN_MAX(struct name *head, int val) { name##_SPLAY_MINMAX(head, val); return (SPLAY_ROOT(head)); }
'' TODO: #define SPLAY_GENERATE(name, type, field, cmp) struct type * name##_SPLAY_INSERT(struct name *head, struct type *elm) { if (SPLAY_EMPTY(head)) { SPLAY_LEFT(elm, field) = SPLAY_RIGHT(elm, field) = NULL; } else { int __comp; name##_SPLAY(head, elm); __comp = (cmp)(elm, (head)->sph_root); if(__comp < 0) { SPLAY_LEFT(elm, field) = SPLAY_LEFT((head)->sph_root, field); SPLAY_RIGHT(elm, field) = (head)->sph_root; SPLAY_LEFT((head)->sph_root, field) = NULL; } else if (__comp > 0) { SPLAY_RIGHT(elm, field) = SPLAY_RIGHT((head)->sph_root, field); SPLAY_LEFT(elm, field) = (head)->sph_root; SPLAY_RIGHT((head)->sph_root, field) = NULL; } else return ((head)->sph_root); } (head)->sph_root = (elm); return (NULL); } struct type * name##_SPLAY_REMOVE(struct name *head, struct type *elm) { struct type *__tmp; if (SPLAY_EMPTY(head)) return (NULL); name##_SPLAY(head, elm); if ((cmp)(elm, (head)->sph_root) == 0) { if (SPLAY_LEFT((head)->sph_root, field) == NULL) { (head)->sph_root = SPLAY_RIGHT((head)->sph_root, field); } else { __tmp = SPLAY_RIGHT((head)->sph_root, field); (head)->sph_root = SPLAY_LEFT((head)->sph_root, field); name##_SPLAY(head, elm); SPLAY_RIGHT((head)->sph_root, field) = __tmp; } return (elm); } return (NULL); } void name##_SPLAY(struct name *head, struct type *elm) { struct type __node, *__left, *__right, *__tmp; int __comp; SPLAY_LEFT(&__node, field) = SPLAY_RIGHT(&__node, field) = NULL; __left = __right = &__node; while ((__comp = (cmp)(elm, (head)->sph_root)) != 0) { if (__comp < 0) { __tmp = SPLAY_LEFT((head)->sph_root, field); if (__tmp == NULL) break; if ((cmp)(elm, __tmp) < 0){ SPLAY_ROTATE_RIGHT(head, __tmp, field); if (SPLAY_LEFT((head)->sph_root, field) == NULL) break; } SPLAY_LINKLEFT(head, __right, field); } else if (__comp > 0) { __tmp = SPLAY_RIGHT((head)->sph_root, field); if (__tmp == NULL) break; if ((cmp)(elm, __tmp) > 0){ SPLAY_ROTATE_LEFT(head, __tmp, field); if (SPLAY_RIGHT((head)->sph_root, field) == NULL) break; } SPLAY_LINKRIGHT(head, __left, field); } } SPLAY_ASSEMBLE(head, &__node, __left, __right, field); } void name##_SPLAY_MINMAX(struct name *head, int __comp) { struct type __node, *__left, *__right, *__tmp; SPLAY_LEFT(&__node, field) = SPLAY_RIGHT(&__node, field) = NULL; __left = __right = &__node; while (1) { if (__comp < 0) { __tmp = SPLAY_LEFT((head)->sph_root, field); if (__tmp == NULL) break; if (__comp < 0){ SPLAY_ROTATE_RIGHT(head, __tmp, field); if (SPLAY_LEFT((head)->sph_root, field) == NULL) break; } SPLAY_LINKLEFT(head, __right, field); } else if (__comp > 0) { __tmp = SPLAY_RIGHT((head)->sph_root, field); if (__tmp == NULL) break; if (__comp > 0) { SPLAY_ROTATE_LEFT(head, __tmp, field); if (SPLAY_RIGHT((head)->sph_root, field) == NULL) break; } SPLAY_LINKRIGHT(head, __left, field); } } SPLAY_ASSEMBLE(head, &__node, __left, __right, field); }
const SPLAY_NEGINF = -1
const SPLAY_INF = 1

'' TODO: #define SPLAY_INSERT(name, x, y) name##_SPLAY_INSERT(x, y)
'' TODO: #define SPLAY_REMOVE(name, x, y) name##_SPLAY_REMOVE(x, y)
'' TODO: #define SPLAY_FIND(name, x, y) name##_SPLAY_FIND(x, y)
'' TODO: #define SPLAY_NEXT(name, x, y) name##_SPLAY_NEXT(x, y)
'' TODO: #define SPLAY_MIN(name, x) (SPLAY_EMPTY(x) ? NULL : name##_SPLAY_MIN_MAX(x, SPLAY_NEGINF))
'' TODO: #define SPLAY_MAX(name, x) (SPLAY_EMPTY(x) ? NULL : name##_SPLAY_MIN_MAX(x, SPLAY_INF))
'' TODO: #define SPLAY_FOREACH(x, name, head) for ((x) = SPLAY_MIN(name, head); (x) != NULL; (x) = SPLAY_NEXT(name, head, x))
'' TODO: #define RB_HEAD(name, type) struct name { struct type *rbh_root; }

#define RB_INITIALIZER(root) (NULL)
#define RB_INIT(root) scope : (root)->rbh_root = NULL : end scope
const RB_BLACK = 0
const RB_RED = 1
'' TODO: #define RB_ENTRY(type) struct { struct type *rbe_left; struct type *rbe_right; struct type *rbe_parent; int rbe_color; }
#define RB_LEFT(elm, field) (elm)->field.rbe_left
#define RB_RIGHT(elm, field) (elm)->field.rbe_right
#define RB_PARENT(elm, field) (elm)->field.rbe_parent
#define RB_COLOR(elm, field) (elm)->field.rbe_color
#define RB_ROOT(head) (head)->rbh_root
#define RB_EMPTY(head) (RB_ROOT(head) = NULL)
#macro RB_SET(elm, parent, field)
	scope
		RB_PARENT(elm, field) = parent
		'' TODO: RB_LEFT(elm, field) = RB_RIGHT(elm, field) = NULL;
		RB_COLOR(elm, field) = RB_RED
	end scope
#endmacro
#macro RB_SET_BLACKRED(black, red, field)
	scope
		RB_COLOR(black, field) = RB_BLACK
		RB_COLOR(red, field) = RB_RED
	end scope
#endmacro
#macro RB_AUGMENT(x)
	scope
	end scope
#endmacro
#macro RB_ROTATE_LEFT(head, elm, tmp, field)
	scope
		(tmp) = RB_RIGHT(elm, field)
		'' TODO: if ((RB_RIGHT(elm, field) = RB_LEFT(tmp, field)) != NULL) { RB_PARENT(RB_LEFT(tmp, field), field) = (elm); }
		RB_AUGMENT(elm)
		'' TODO: if ((RB_PARENT(tmp, field) = RB_PARENT(elm, field)) != NULL) { if ((elm) == RB_LEFT(RB_PARENT(elm, field), field)) RB_LEFT(RB_PARENT(elm, field), field) = (tmp); else RB_RIGHT(RB_PARENT(elm, field), field) = (tmp); }
		'' TODO: else (head)->rbh_root = (tmp);
		RB_LEFT(tmp, field) = (elm)
		RB_PARENT(elm, field) = (tmp)
		RB_AUGMENT(tmp)
		if RB_PARENT(tmp, field) then
			RB_AUGMENT(RB_PARENT(tmp, field))
		end if
	end scope
#endmacro
#macro RB_ROTATE_RIGHT(head, elm, tmp, field)
	scope
		(tmp) = RB_LEFT(elm, field)
		'' TODO: if ((RB_LEFT(elm, field) = RB_RIGHT(tmp, field)) != NULL) { RB_PARENT(RB_RIGHT(tmp, field), field) = (elm); }
		RB_AUGMENT(elm)
		'' TODO: if ((RB_PARENT(tmp, field) = RB_PARENT(elm, field)) != NULL) { if ((elm) == RB_LEFT(RB_PARENT(elm, field), field)) RB_LEFT(RB_PARENT(elm, field), field) = (tmp); else RB_RIGHT(RB_PARENT(elm, field), field) = (tmp); }
		'' TODO: else (head)->rbh_root = (tmp);
		RB_RIGHT(tmp, field) = (elm)
		RB_PARENT(elm, field) = (tmp)
		RB_AUGMENT(tmp)
		if RB_PARENT(tmp, field) then
			RB_AUGMENT(RB_PARENT(tmp, field))
		end if
	end scope
#endmacro

'' TODO: #define RB_PROTOTYPE(name, type, field, cmp) RB_PROTOTYPE_INTERNAL(name, type, field, cmp,)
'' TODO: #define RB_PROTOTYPE_STATIC(name, type, field, cmp) RB_PROTOTYPE_INTERNAL(name, type, field, cmp, __unused static)
'' TODO: #define RB_PROTOTYPE_INTERNAL(name, type, field, cmp, attr) RB_PROTOTYPE_INSERT_COLOR(name, type, attr); RB_PROTOTYPE_REMOVE_COLOR(name, type, attr); RB_PROTOTYPE_INSERT(name, type, attr); RB_PROTOTYPE_REMOVE(name, type, attr); RB_PROTOTYPE_FIND(name, type, attr); RB_PROTOTYPE_NFIND(name, type, attr); RB_PROTOTYPE_NEXT(name, type, attr); RB_PROTOTYPE_PREV(name, type, attr); RB_PROTOTYPE_MINMAX(name, type, attr);
'' TODO: #define RB_PROTOTYPE_INSERT_COLOR(name, type, attr) attr void name##_RB_INSERT_COLOR(struct name *, struct type *)
'' TODO: #define RB_PROTOTYPE_REMOVE_COLOR(name, type, attr) attr void name##_RB_REMOVE_COLOR(struct name *, struct type *, struct type *)
'' TODO: #define RB_PROTOTYPE_REMOVE(name, type, attr) attr struct type *name##_RB_REMOVE(struct name *, struct type *)
'' TODO: #define RB_PROTOTYPE_INSERT(name, type, attr) attr struct type *name##_RB_INSERT(struct name *, struct type *)
'' TODO: #define RB_PROTOTYPE_FIND(name, type, attr) attr struct type *name##_RB_FIND(struct name *, struct type *)
'' TODO: #define RB_PROTOTYPE_NFIND(name, type, attr) attr struct type *name##_RB_NFIND(struct name *, struct type *)
'' TODO: #define RB_PROTOTYPE_NEXT(name, type, attr) attr struct type *name##_RB_NEXT(struct type *)
'' TODO: #define RB_PROTOTYPE_PREV(name, type, attr) attr struct type *name##_RB_PREV(struct type *)
'' TODO: #define RB_PROTOTYPE_MINMAX(name, type, attr) attr struct type *name##_RB_MINMAX(struct name *, int)
'' TODO: #define RB_GENERATE(name, type, field, cmp) RB_GENERATE_INTERNAL(name, type, field, cmp,)
'' TODO: #define RB_GENERATE_STATIC(name, type, field, cmp) RB_GENERATE_INTERNAL(name, type, field, cmp, __unused static)
'' TODO: #define RB_GENERATE_INTERNAL(name, type, field, cmp, attr) RB_GENERATE_INSERT_COLOR(name, type, field, attr) RB_GENERATE_REMOVE_COLOR(name, type, field, attr) RB_GENERATE_INSERT(name, type, field, cmp, attr) RB_GENERATE_REMOVE(name, type, field, attr) RB_GENERATE_FIND(name, type, field, cmp, attr) RB_GENERATE_NFIND(name, type, field, cmp, attr) RB_GENERATE_NEXT(name, type, field, attr) RB_GENERATE_PREV(name, type, field, attr) RB_GENERATE_MINMAX(name, type, field, attr)
'' TODO: #define RB_GENERATE_INSERT_COLOR(name, type, field, attr) attr void name##_RB_INSERT_COLOR(struct name *head, struct type *elm) { struct type *parent, *gparent, *tmp; while ((parent = RB_PARENT(elm, field)) != NULL && RB_COLOR(parent, field) == RB_RED) { gparent = RB_PARENT(parent, field); if (parent == RB_LEFT(gparent, field)) { tmp = RB_RIGHT(gparent, field); if (tmp && RB_COLOR(tmp, field) == RB_RED) { RB_COLOR(tmp, field) = RB_BLACK; RB_SET_BLACKRED(parent, gparent, field); elm = gparent; continue; } if (RB_RIGHT(parent, field) == elm) { RB_ROTATE_LEFT(head, parent, tmp, field); tmp = parent; parent = elm; elm = tmp; } RB_SET_BLACKRED(parent, gparent, field); RB_ROTATE_RIGHT(head, gparent, tmp, field); } else { tmp = RB_LEFT(gparent, field); if (tmp && RB_COLOR(tmp, field) == RB_RED) { RB_COLOR(tmp, field) = RB_BLACK; RB_SET_BLACKRED(parent, gparent, field); elm = gparent; continue; } if (RB_LEFT(parent, field) == elm) { RB_ROTATE_RIGHT(head, parent, tmp, field); tmp = parent; parent = elm; elm = tmp; } RB_SET_BLACKRED(parent, gparent, field); RB_ROTATE_LEFT(head, gparent, tmp, field); } } RB_COLOR(head->rbh_root, field) = RB_BLACK; }
'' TODO: #define RB_GENERATE_REMOVE_COLOR(name, type, field, attr) attr void name##_RB_REMOVE_COLOR(struct name *head, struct type *parent, struct type *elm) { struct type *tmp; while ((elm == NULL || RB_COLOR(elm, field) == RB_BLACK) && elm != RB_ROOT(head)) { if (RB_LEFT(parent, field) == elm) { tmp = RB_RIGHT(parent, field); if (RB_COLOR(tmp, field) == RB_RED) { RB_SET_BLACKRED(tmp, parent, field); RB_ROTATE_LEFT(head, parent, tmp, field); tmp = RB_RIGHT(parent, field); } if ((RB_LEFT(tmp, field) == NULL || RB_COLOR(RB_LEFT(tmp, field), field) == RB_BLACK) && (RB_RIGHT(tmp, field) == NULL || RB_COLOR(RB_RIGHT(tmp, field), field) == RB_BLACK)) { RB_COLOR(tmp, field) = RB_RED; elm = parent; parent = RB_PARENT(elm, field); } else { if (RB_RIGHT(tmp, field) == NULL || RB_COLOR(RB_RIGHT(tmp, field), field) == RB_BLACK) { struct type *oleft; if ((oleft = RB_LEFT(tmp, field)) != NULL) RB_COLOR(oleft, field) = RB_BLACK; RB_COLOR(tmp, field) = RB_RED; RB_ROTATE_RIGHT(head, tmp, oleft, field); tmp = RB_RIGHT(parent, field); } RB_COLOR(tmp, field) = RB_COLOR(parent, field); RB_COLOR(parent, field) = RB_BLACK; if (RB_RIGHT(tmp, field)) RB_COLOR(RB_RIGHT(tmp, field), field) = RB_BLACK; RB_ROTATE_LEFT(head, parent, tmp, field); elm = RB_ROOT(head); break; } } else { tmp = RB_LEFT(parent, field); if (RB_COLOR(tmp, field) == RB_RED) { RB_SET_BLACKRED(tmp, parent, field); RB_ROTATE_RIGHT(head, parent, tmp, field); tmp = RB_LEFT(parent, field); } if ((RB_LEFT(tmp, field) == NULL || RB_COLOR(RB_LEFT(tmp, field), field) == RB_BLACK) && (RB_RIGHT(tmp, field) == NULL || RB_COLOR(RB_RIGHT(tmp, field), field) == RB_BLACK)) { RB_COLOR(tmp, field) = RB_RED; elm = parent; parent = RB_PARENT(elm, field); } else { if (RB_LEFT(tmp, field) == NULL || RB_COLOR(RB_LEFT(tmp, field), field) == RB_BLACK) { struct type *oright; if ((oright = RB_RIGHT(tmp, field)) != NULL) RB_COLOR(oright, field) = RB_BLACK; RB_COLOR(tmp, field) = RB_RED; RB_ROTATE_LEFT(head, tmp, oright, field); tmp = RB_LEFT(parent, field); } RB_COLOR(tmp, field) = RB_COLOR(parent, field); RB_COLOR(parent, field) = RB_BLACK; if (RB_LEFT(tmp, field)) RB_COLOR(RB_LEFT(tmp, field), field) = RB_BLACK; RB_ROTATE_RIGHT(head, parent, tmp, field); elm = RB_ROOT(head); break; } } } if (elm) RB_COLOR(elm, field) = RB_BLACK; }
'' TODO: #define RB_GENERATE_REMOVE(name, type, field, attr) attr struct type * name##_RB_REMOVE(struct name *head, struct type *elm) { struct type *child, *parent, *old = elm; int color; if (RB_LEFT(elm, field) == NULL) child = RB_RIGHT(elm, field); else if (RB_RIGHT(elm, field) == NULL) child = RB_LEFT(elm, field); else { struct type *left; elm = RB_RIGHT(elm, field); while ((left = RB_LEFT(elm, field)) != NULL) elm = left; child = RB_RIGHT(elm, field); parent = RB_PARENT(elm, field); color = RB_COLOR(elm, field); if (child) RB_PARENT(child, field) = parent; if (parent) { if (RB_LEFT(parent, field) == elm) RB_LEFT(parent, field) = child; else RB_RIGHT(parent, field) = child; RB_AUGMENT(parent); } else RB_ROOT(head) = child; if (RB_PARENT(elm, field) == old) parent = elm; (elm)->field = (old)->field; if (RB_PARENT(old, field)) { if (RB_LEFT(RB_PARENT(old, field), field) == old) RB_LEFT(RB_PARENT(old, field), field) = elm; else RB_RIGHT(RB_PARENT(old, field), field) = elm; RB_AUGMENT(RB_PARENT(old, field)); } else RB_ROOT(head) = elm; RB_PARENT(RB_LEFT(old, field), field) = elm; if (RB_RIGHT(old, field)) RB_PARENT(RB_RIGHT(old, field), field) = elm; if (parent) { left = parent; do { RB_AUGMENT(left); } while ((left = RB_PARENT(left, field)) != NULL); } goto color; } parent = RB_PARENT(elm, field); color = RB_COLOR(elm, field); if (child) RB_PARENT(child, field) = parent; if (parent) { if (RB_LEFT(parent, field) == elm) RB_LEFT(parent, field) = child; else RB_RIGHT(parent, field) = child; RB_AUGMENT(parent); } else RB_ROOT(head) = child; color: if (color == RB_BLACK) name##_RB_REMOVE_COLOR(head, parent, child); return (old); }
'' TODO: #define RB_GENERATE_INSERT(name, type, field, cmp, attr) attr struct type * name##_RB_INSERT(struct name *head, struct type *elm) { struct type *tmp; struct type *parent = NULL; int comp = 0; tmp = RB_ROOT(head); while (tmp) { parent = tmp; comp = (cmp)(elm, parent); if (comp < 0) tmp = RB_LEFT(tmp, field); else if (comp > 0) tmp = RB_RIGHT(tmp, field); else return (tmp); } RB_SET(elm, parent, field); if (parent != NULL) { if (comp < 0) RB_LEFT(parent, field) = elm; else RB_RIGHT(parent, field) = elm; RB_AUGMENT(parent); } else RB_ROOT(head) = elm; name##_RB_INSERT_COLOR(head, elm); return (NULL); }
'' TODO: #define RB_GENERATE_FIND(name, type, field, cmp, attr) attr struct type * name##_RB_FIND(struct name *head, struct type *elm) { struct type *tmp = RB_ROOT(head); int comp; while (tmp) { comp = cmp(elm, tmp); if (comp < 0) tmp = RB_LEFT(tmp, field); else if (comp > 0) tmp = RB_RIGHT(tmp, field); else return (tmp); } return (NULL); }
'' TODO: #define RB_GENERATE_NFIND(name, type, field, cmp, attr) attr struct type * name##_RB_NFIND(struct name *head, struct type *elm) { struct type *tmp = RB_ROOT(head); struct type *res = NULL; int comp; while (tmp) { comp = cmp(elm, tmp); if (comp < 0) { res = tmp; tmp = RB_LEFT(tmp, field); } else if (comp > 0) tmp = RB_RIGHT(tmp, field); else return (tmp); } return (res); }
'' TODO: #define RB_GENERATE_NEXT(name, type, field, attr) attr struct type * name##_RB_NEXT(struct type *elm) { if (RB_RIGHT(elm, field)) { elm = RB_RIGHT(elm, field); while (RB_LEFT(elm, field)) elm = RB_LEFT(elm, field); } else { if (RB_PARENT(elm, field) && (elm == RB_LEFT(RB_PARENT(elm, field), field))) elm = RB_PARENT(elm, field); else { while (RB_PARENT(elm, field) && (elm == RB_RIGHT(RB_PARENT(elm, field), field))) elm = RB_PARENT(elm, field); elm = RB_PARENT(elm, field); } } return (elm); }
'' TODO: #define RB_GENERATE_PREV(name, type, field, attr) attr struct type * name##_RB_PREV(struct type *elm) { if (RB_LEFT(elm, field)) { elm = RB_LEFT(elm, field); while (RB_RIGHT(elm, field)) elm = RB_RIGHT(elm, field); } else { if (RB_PARENT(elm, field) && (elm == RB_RIGHT(RB_PARENT(elm, field), field))) elm = RB_PARENT(elm, field); else { while (RB_PARENT(elm, field) && (elm == RB_LEFT(RB_PARENT(elm, field), field))) elm = RB_PARENT(elm, field); elm = RB_PARENT(elm, field); } } return (elm); }
'' TODO: #define RB_GENERATE_MINMAX(name, type, field, attr) attr struct type * name##_RB_MINMAX(struct name *head, int val) { struct type *tmp = RB_ROOT(head); struct type *parent = NULL; while (tmp) { parent = tmp; if (val < 0) tmp = RB_LEFT(tmp, field); else tmp = RB_RIGHT(tmp, field); } return (parent); }
const RB_NEGINF = -1
const RB_INF = 1
'' TODO: #define RB_INSERT(name, x, y) name##_RB_INSERT(x, y)
'' TODO: #define RB_REMOVE(name, x, y) name##_RB_REMOVE(x, y)
'' TODO: #define RB_FIND(name, x, y) name##_RB_FIND(x, y)
'' TODO: #define RB_NFIND(name, x, y) name##_RB_NFIND(x, y)
'' TODO: #define RB_NEXT(name, x, y) name##_RB_NEXT(y)
'' TODO: #define RB_PREV(name, x, y) name##_RB_PREV(y)
'' TODO: #define RB_MIN(name, x) name##_RB_MINMAX(x, RB_NEGINF)
'' TODO: #define RB_MAX(name, x) name##_RB_MINMAX(x, RB_INF)
'' TODO: #define RB_FOREACH(x, name, head) for ((x) = RB_MIN(name, head); (x) != NULL; (x) = name##_RB_NEXT(x))
'' TODO: #define RB_FOREACH_FROM(x, name, y) for ((x) = (y); ((x) != NULL) && ((y) = name##_RB_NEXT(x), (x) != NULL); (x) = (y))
'' TODO: #define RB_FOREACH_SAFE(x, name, head, y) for ((x) = RB_MIN(name, head); ((x) != NULL) && ((y) = name##_RB_NEXT(x), (x) != NULL); (x) = (y))
'' TODO: #define RB_FOREACH_REVERSE(x, name, head) for ((x) = RB_MAX(name, head); (x) != NULL; (x) = name##_RB_PREV(x))
'' TODO: #define RB_FOREACH_REVERSE_FROM(x, name, y) for ((x) = (y); ((x) != NULL) && ((y) = name##_RB_PREV(x), (x) != NULL); (x) = (y))
'' TODO: #define RB_FOREACH_REVERSE_SAFE(x, name, head, y) for ((x) = RB_MAX(name, head); ((x) != NULL) && ((y) = name##_RB_PREV(x), (x) != NULL); (x) = (y))
'' TODO: #define STATIC static
type lxw_row_t as ulong
type lxw_col_t as ushort

type lxw_boolean as long
enum
	LXW_FALSE
	LXW_TRUE
end enum

type lxw_error as long
enum
	LXW_NO_ERROR = 0
	LXW_ERROR_MEMORY_MALLOC_FAILED
	LXW_ERROR_CREATING_XLSX_FILE
	LXW_ERROR_CREATING_TMPFILE
	LXW_ERROR_ZIP_FILE_OPERATION
	LXW_ERROR_ZIP_FILE_ADD
	LXW_ERROR_ZIP_CLOSE
	LXW_ERROR_NULL_PARAMETER_IGNORED
	LXW_ERROR_PARAMETER_VALIDATION
	LXW_ERROR_SHEETNAME_LENGTH_EXCEEDED
	LXW_ERROR_INVALID_SHEETNAME_CHARACTER
	LXW_ERROR_SHEETNAME_ALREADY_USED
	LXW_ERROR_128_STRING_LENGTH_EXCEEDED
	LXW_ERROR_255_STRING_LENGTH_EXCEEDED
	LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED
	LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND
	LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE
	LXW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED
	LXW_ERROR_IMAGE_DIMENSIONS
	LXW_MAX_ERRNO
end enum

type lxw_datetime
	year as long
	month as long
	day as long
	hour as long
	min as long
	sec as double
end type

type lxw_custom_property_types as long
enum
	LXW_CUSTOM_NONE
	LXW_CUSTOM_STRING
	LXW_CUSTOM_DOUBLE
	LXW_CUSTOM_INTEGER
	LXW_CUSTOM_BOOLEAN
	LXW_CUSTOM_DATETIME
end enum

const LXW_SHEETNAME_MAX = 31
const LXW_MAX_SHEETNAME_LENGTH = ((LXW_SHEETNAME_MAX * 4) + 2) + 1
#define LXW_MAX_COL_NAME_LENGTH sizeof("$XFD")
#define LXW_MAX_ROW_NAME_LENGTH sizeof("$1048576")
#define LXW_MAX_CELL_NAME_LENGTH sizeof("$XFWD$1048576")
#define LXW_MAX_CELL_RANGE_LENGTH (LXW_MAX_CELL_NAME_LENGTH * 2)
#define LXW_MAX_FORMULA_RANGE_LENGTH (LXW_MAX_SHEETNAME_LENGTH + LXW_MAX_CELL_RANGE_LENGTH)
#define LXW_DATETIME_LENGTH sizeof("2016-12-12T23:00:00Z")
const LXW_EPOCH_1900 = 0
const LXW_EPOCH_1904 = 1
#define LXW_UINT32_T_LENGTH sizeof("4294967296")
const LXW_FILENAME_LENGTH = 128
const LXW_IGNORE = 1
#define LXW_SCHEMA_MS "http://schemas.microsoft.com/office/2006/relationships"
#define LXW_SCHEMA_ROOT "http://schemas.openxmlformats.org"
#define LXW_SCHEMA_DRAWING LXW_SCHEMA_ROOT "/drawingml/2006"
#define LXW_SCHEMA_OFFICEDOC LXW_SCHEMA_ROOT "/officeDocument/2006"
#define LXW_SCHEMA_PACKAGE LXW_SCHEMA_ROOT "/package/2006/relationships"
#define LXW_SCHEMA_DOCUMENT LXW_SCHEMA_ROOT "/officeDocument/2006/relationships"
#define LXW_SCHEMA_CONTENT LXW_SCHEMA_ROOT "/package/2006/content-types"
' ~ #define LXW_ERROR(message) fprintf(stderr, "[ERROR][%s:%d]: " message !"\n", __FILE__, __LINE__)
#define LXW_MEM_ERROR() LXW_ERROR("Memory allocation failed.")
#macro GOTO_LABEL_ON_MEM_ERROR(pointer, label)
	if pointer = 0 then
		LXW_MEM_ERROR()
		'' TODO: goto label;
	end if
#endmacro
#macro RETURN_ON_MEM_ERROR(pointer, error)
	if pointer = 0 then
		LXW_MEM_ERROR()
		return error
	end if
#endmacro
#macro RETURN_VOID_ON_MEM_ERROR(pointer)
	if pointer = 0 then
		LXW_MEM_ERROR()
		return
	end if
#endmacro
#macro RETURN_ON_ERROR(error)
	if error then
		return error
	end if
#endmacro
#define LXW_WARN(message) fprintf(stderr, "[WARNING]: " message !"\n")
#define LXW_WARN_FORMAT(message) fprintf(stderr, "[WARNING]: " message !"\n")
#define LXW_WARN_FORMAT1(message, var) fprintf(stderr, "[WARNING]: " message !"\n", var)
#define LXW_WARN_FORMAT2(message, var1, var2) fprintf(stderr, "[WARNING]: " message !"\n", var1, var2)
#macro LXW_WARN_CAT_AXIS_ONLY(function)
	if axis->is_category = 0 then
		fprintf(stderr, "[WARNING]: " function !"() is only valid for category axes\n")
		return
	end if
#endmacro
#macro LXW_WARN_VALUE_AXIS_ONLY(function)
	if axis->is_value = 0 then
		fprintf(stderr, "[WARNING]: " function !"() is only valid for value axes\n")
		return
	end if
#endmacro
#macro LXW_WARN_DATE_AXIS_ONLY(function)
	if axis->is_date = 0 then
		fprintf(stderr, "[WARNING]: " function !"() is only valid for date axes\n")
		return
	end if
#endmacro
#macro LXW_WARN_CAT_AND_DATE_AXIS_ONLY(function)
	if (axis->is_category = 0) andalso (axis->is_date = 0) then
		fprintf(stderr, "[WARNING]: " function !"() is only valid for category and date axes\n")
		return
	end if
#endmacro
#macro LXW_WARN_VALUE_AND_DATE_AXIS_ONLY(function)
	if (axis->is_value = 0) andalso (axis->is_date = 0) then
		fprintf(stderr, "[WARNING]: " function !"() is only valid for value and date axes\n")
		return
	end if
#endmacro
#define LXW_UINT32_NETWORK(n) ((((((n) and &hFF) shl 24) or (((n) and &hFF00) shl 8)) or (((n) and &hFF0000) shr 8)) or (((n) and &hFF000000) shr 24))
#define LXW_UINT16_NETWORK(n) ((((n) and &h00FF) shl 8) or (((n) and &hFF00) shr 8))
#define lxw_snprintf __builtin_snprintf
#define lxw_strcpy(dest, src) lxw_snprintf(dest, sizeof(dest), "%s", src)
type lxw_format as lxw_format_

type lxw_formats
	stqh_first as lxw_format ptr
	stqh_last as lxw_format ptr ptr
end type

type lxw_tuple as lxw_tuple_

type lxw_tuples
	stqh_first as lxw_tuple ptr
	stqh_last as lxw_tuple ptr ptr
end type

type lxw_custom_property as lxw_custom_property_

type lxw_custom_properties
	stqh_first as lxw_custom_property ptr
	stqh_last as lxw_custom_property ptr ptr
end type

type lxw_tuple_list_pointers
	stqe_next as lxw_tuple ptr
end type

type lxw_tuple_
	key as zstring ptr
	value as zstring ptr
	list_pointers as lxw_tuple_list_pointers
end type

union lxw_custom_property_u
	string as zstring ptr
	number as double
	integer as long
	boolean as ubyte
	datetime as lxw_datetime
end union

type lxw_custom_property_list_pointers
	stqe_next as lxw_custom_property ptr
end type

type lxw_custom_property_
	as lxw_custom_property_types type
	name as zstring ptr
	u as lxw_custom_property_u
	list_pointers as lxw_custom_property_list_pointers
end type

type sst_element as sst_element_

type sst_rb_tree
	rbh_root as sst_element ptr
end type

type sst_order_list
	stqh_first as sst_element ptr
	stqh_last as sst_element ptr ptr
end type

'' TODO: #define LXW_RB_GENERATE_ELEMENT(name, type, field, cmp) RB_GENERATE_INSERT_COLOR(name, type, field, static) RB_GENERATE_INSERT(name, type, field, cmp, static) struct lxw_rb_generate_element{int unused;}

type sst_element_sst_order_pointers
	stqe_next as sst_element ptr
end type

type sst_element_sst_tree_pointers
	rbe_left as sst_element ptr
	rbe_right as sst_element ptr
	rbe_parent as sst_element ptr
	rbe_color as long
end type

type sst_element_
	index as ulong
	string as zstring ptr
	sst_order_pointers as sst_element_sst_order_pointers
	sst_tree_pointers as sst_element_sst_tree_pointers
end type

type lxw_sst
	file as FILE ptr
	string_count as ulong
	unique_count as ulong
	order_list as sst_order_list ptr
	rb_tree as sst_rb_tree ptr
end type

declare function lxw_sst_new() as lxw_sst ptr
declare sub lxw_sst_free(byval sst as lxw_sst ptr)
declare function lxw_get_sst_index(byval sst as lxw_sst ptr, byval string as const zstring ptr) as sst_element ptr
declare sub lxw_sst_assemble_xml_file(byval self as lxw_sst ptr)

#define __LXW_CHART_H__
#define __LXW_FORMAT_H__
#define __LXW_HASH_TABLE_H__
#define LXW_FOREACH_ORDERED(elem, hash_table) STAILQ_FOREACH((elem), (hash_table)->order_list, lxw_hash_order_pointers)
type lxw_hash_element as lxw_hash_element_

type lxw_hash_order_list
	stqh_first as lxw_hash_element ptr
	stqh_last as lxw_hash_element ptr ptr
end type

type lxw_hash_bucket_list
	slh_first as lxw_hash_element ptr
end type

type lxw_hash_table
	num_buckets as ulong
	used_buckets as ulong
	unique_count as ulong
	free_key as ubyte
	free_value as ubyte
	order_list as lxw_hash_order_list ptr
	buckets as lxw_hash_bucket_list ptr ptr
end type

type lxw_hash_element_lxw_hash_order_pointers
	stqe_next as lxw_hash_element ptr
end type

type lxw_hash_element_lxw_hash_list_pointers
	sle_next as lxw_hash_element ptr
end type

type lxw_hash_element_
	key as any ptr
	value as any ptr
	lxw_hash_order_pointers as lxw_hash_element_lxw_hash_order_pointers
	lxw_hash_list_pointers as lxw_hash_element_lxw_hash_list_pointers
end type

declare function lxw_hash_key_exists(byval lxw_hash as lxw_hash_table ptr, byval key as any ptr, byval key_len as uinteger) as lxw_hash_element ptr
declare function lxw_insert_hash_element(byval lxw_hash as lxw_hash_table ptr, byval key as any ptr, byval value as any ptr, byval key_len as uinteger) as lxw_hash_element ptr
declare function lxw_hash_new(byval num_buckets as ulong, byval free_key as ubyte, byval free_value as ubyte) as lxw_hash_table ptr
declare sub lxw_hash_free(byval lxw_hash as lxw_hash_table ptr)
type lxw_color_t as long

const LXW_FORMAT_FIELD_LEN = 128
#define LXW_DEFAULT_FONT_NAME "Calibri"
const LXW_DEFAULT_FONT_FAMILY = 2
const LXW_DEFAULT_FONT_THEME = 1
const LXW_PROPERTY_UNSET = -1
const LXW_COLOR_UNSET = -1
const LXW_COLOR_MASK = &hFFFFFF
const LXW_MIN_FONT_SIZE = 1
const LXW_MAX_FONT_SIZE = 409
#macro LXW_FORMAT_FIELD_COPY(dst, src)
	scope
		strncpy(dst, src, LXW_FORMAT_FIELD_LEN - 1)
		dst[(LXW_FORMAT_FIELD_LEN - 1)] = asc(!"\0")
	end scope
#endmacro

type lxw_format_underlines as long
enum
	LXW_UNDERLINE_SINGLE = 1
	LXW_UNDERLINE_DOUBLE
	LXW_UNDERLINE_SINGLE_ACCOUNTING
	LXW_UNDERLINE_DOUBLE_ACCOUNTING
end enum

type lxw_format_scripts as long
enum
	LXW_FONT_SUPERSCRIPT = 1
	LXW_FONT_SUBSCRIPT
end enum

type lxw_format_alignments as long
enum
	LXW_ALIGN_NONE = 0
	LXW_ALIGN_LEFT
	LXW_ALIGN_CENTER
	LXW_ALIGN_RIGHT
	LXW_ALIGN_FILL
	LXW_ALIGN_JUSTIFY
	LXW_ALIGN_CENTER_ACROSS
	LXW_ALIGN_DISTRIBUTED
	LXW_ALIGN_VERTICAL_TOP
	LXW_ALIGN_VERTICAL_BOTTOM
	LXW_ALIGN_VERTICAL_CENTER
	LXW_ALIGN_VERTICAL_JUSTIFY
	LXW_ALIGN_VERTICAL_DISTRIBUTED
end enum

type lxw_format_diagonal_types as long
enum
	LXW_DIAGONAL_BORDER_UP = 1
	LXW_DIAGONAL_BORDER_DOWN
	LXW_DIAGONAL_BORDER_UP_DOWN
end enum

type lxw_defined_colors as long
enum
	LXW_COLOR_BLACK = &h1000000
	LXW_COLOR_BLUE = &h0000FF
	LXW_COLOR_BROWN = &h800000
	LXW_COLOR_CYAN = &h00FFFF
	LXW_COLOR_GRAY = &h808080
	LXW_COLOR_GREEN = &h008000
	LXW_COLOR_LIME = &h00FF00
	LXW_COLOR_MAGENTA = &hFF00FF
	LXW_COLOR_NAVY = &h000080
	LXW_COLOR_ORANGE = &hFF6600
	LXW_COLOR_PINK = &hFF00FF
	LXW_COLOR_PURPLE = &h800080
	LXW_COLOR_RED = &hFF0000
	LXW_COLOR_SILVER = &hC0C0C0
	LXW_COLOR_WHITE = &hFFFFFF
	LXW_COLOR_YELLOW = &hFFFF00
end enum

type lxw_format_patterns as long
enum
	LXW_PATTERN_NONE = 0
	LXW_PATTERN_SOLID
	LXW_PATTERN_MEDIUM_GRAY
	LXW_PATTERN_DARK_GRAY
	LXW_PATTERN_LIGHT_GRAY
	LXW_PATTERN_DARK_HORIZONTAL
	LXW_PATTERN_DARK_VERTICAL
	LXW_PATTERN_DARK_DOWN
	LXW_PATTERN_DARK_UP
	LXW_PATTERN_DARK_GRID
	LXW_PATTERN_DARK_TRELLIS
	LXW_PATTERN_LIGHT_HORIZONTAL
	LXW_PATTERN_LIGHT_VERTICAL
	LXW_PATTERN_LIGHT_DOWN
	LXW_PATTERN_LIGHT_UP
	LXW_PATTERN_LIGHT_GRID
	LXW_PATTERN_LIGHT_TRELLIS
	LXW_PATTERN_GRAY_125
	LXW_PATTERN_GRAY_0625
end enum

type lxw_format_borders as long
enum
	LXW_BORDER_NONE
	LXW_BORDER_THIN
	LXW_BORDER_MEDIUM
	LXW_BORDER_DASHED
	LXW_BORDER_DOTTED
	LXW_BORDER_THICK
	LXW_BORDER_DOUBLE
	LXW_BORDER_HAIR
	LXW_BORDER_MEDIUM_DASHED
	LXW_BORDER_DASH_DOT
	LXW_BORDER_MEDIUM_DASH_DOT
	LXW_BORDER_DASH_DOT_DOT
	LXW_BORDER_MEDIUM_DASH_DOT_DOT
	LXW_BORDER_SLANT_DASH_DOT
end enum

type lxw_format_list_pointers
	stqe_next as lxw_format ptr
end type

type lxw_format_
	file as FILE ptr
	xf_format_indices as lxw_hash_table ptr
	num_xf_formats as ushort ptr
	xf_index as long
	dxf_index as long
	num_format as zstring * 128
	font_name as zstring * 128
	font_scheme as zstring * 128
	num_format_index as ushort
	font_index as ushort
	has_font as ubyte
	has_dxf_font as ubyte
	font_size as ushort
	bold as ubyte
	italic as ubyte
	font_color as lxw_color_t
	underline as ubyte
	font_strikeout as ubyte
	font_outline as ubyte
	font_shadow as ubyte
	font_script as ubyte
	font_family as ubyte
	font_charset as ubyte
	font_condense as ubyte
	font_extend as ubyte
	theme as ubyte
	hyperlink as ubyte
	hidden as ubyte
	locked as ubyte
	text_h_align as ubyte
	text_wrap as ubyte
	text_v_align as ubyte
	text_justlast as ubyte
	rotation as short
	fg_color as lxw_color_t
	bg_color as lxw_color_t
	pattern as ubyte
	has_fill as ubyte
	has_dxf_fill as ubyte
	fill_index as long
	fill_count as long
	border_index as long
	has_border as ubyte
	has_dxf_border as ubyte
	border_count as long
	bottom as ubyte
	diag_border as ubyte
	diag_type as ubyte
	left as ubyte
	right as ubyte
	top as ubyte
	bottom_color as lxw_color_t
	diag_color as lxw_color_t
	left_color as lxw_color_t
	right_color as lxw_color_t
	top_color as lxw_color_t
	indent as ubyte
	shrink as ubyte
	merge_range as ubyte
	reading_order as ubyte
	just_distrib as ubyte
	color_indexed as ubyte
	font_only as ubyte
	list_pointers as lxw_format_list_pointers
end type

type lxw_font
	font_name as zstring * 128
	font_size as ushort
	bold as ubyte
	italic as ubyte
	underline as ubyte
	font_strikeout as ubyte
	font_outline as ubyte
	font_shadow as ubyte
	font_script as ubyte
	font_family as ubyte
	font_charset as ubyte
	font_condense as ubyte
	font_extend as ubyte
	font_color as lxw_color_t
end type

type lxw_border
	bottom as ubyte
	diag_border as ubyte
	diag_type as ubyte
	left as ubyte
	right as ubyte
	top as ubyte
	bottom_color as lxw_color_t
	diag_color as lxw_color_t
	left_color as lxw_color_t
	right_color as lxw_color_t
	top_color as lxw_color_t
end type

type lxw_fill
	fg_color as lxw_color_t
	bg_color as lxw_color_t
	pattern as ubyte
end type

declare function lxw_format_new() as lxw_format ptr
declare sub lxw_format_free(byval format as lxw_format ptr)
declare function lxw_format_get_xf_index(byval format as lxw_format ptr) as long
declare function lxw_format_get_font_key(byval format as lxw_format ptr) as lxw_font ptr
declare function lxw_format_get_border_key(byval format as lxw_format ptr) as lxw_border ptr
declare function lxw_format_get_fill_key(byval format as lxw_format ptr) as lxw_fill ptr
declare function lxw_format_check_color(byval color as lxw_color_t) as lxw_color_t
declare sub format_set_font_name(byval format as lxw_format ptr, byval font_name as const zstring ptr)
declare sub format_set_font_size(byval format as lxw_format ptr, byval size as ushort)
declare sub format_set_font_color(byval format as lxw_format ptr, byval color as lxw_color_t)
declare sub format_set_bold(byval format as lxw_format ptr)
declare sub format_set_italic(byval format as lxw_format ptr)
declare sub format_set_underline(byval format as lxw_format ptr, byval style as ubyte)
declare sub format_set_font_strikeout(byval format as lxw_format ptr)
declare sub format_set_font_script(byval format as lxw_format ptr, byval style as ubyte)
declare sub format_set_num_format(byval format as lxw_format ptr, byval num_format as const zstring ptr)
declare sub format_set_num_format_index(byval format as lxw_format ptr, byval index as ubyte)
declare sub format_set_unlocked(byval format as lxw_format ptr)
declare sub format_set_hidden(byval format as lxw_format ptr)
declare sub format_set_align(byval format as lxw_format ptr, byval alignment as ubyte)
declare sub format_set_text_wrap(byval format as lxw_format ptr)
declare sub format_set_rotation(byval format as lxw_format ptr, byval angle as short)
declare sub format_set_indent(byval format as lxw_format ptr, byval level as ubyte)
declare sub format_set_shrink(byval format as lxw_format ptr)
declare sub format_set_pattern(byval format as lxw_format ptr, byval index as ubyte)
declare sub format_set_bg_color(byval format as lxw_format ptr, byval color as lxw_color_t)
declare sub format_set_fg_color(byval format as lxw_format ptr, byval color as lxw_color_t)
declare sub format_set_border(byval format as lxw_format ptr, byval style as ubyte)
declare sub format_set_bottom(byval format as lxw_format ptr, byval style as ubyte)
declare sub format_set_top(byval format as lxw_format ptr, byval style as ubyte)
declare sub format_set_left(byval format as lxw_format ptr, byval style as ubyte)
declare sub format_set_right(byval format as lxw_format ptr, byval style as ubyte)
declare sub format_set_border_color(byval format as lxw_format ptr, byval color as lxw_color_t)
declare sub format_set_bottom_color(byval format as lxw_format ptr, byval color as lxw_color_t)
declare sub format_set_top_color(byval format as lxw_format ptr, byval color as lxw_color_t)
declare sub format_set_left_color(byval format as lxw_format ptr, byval color as lxw_color_t)
declare sub format_set_right_color(byval format as lxw_format ptr, byval color as lxw_color_t)
declare sub format_set_diag_type(byval format as lxw_format ptr, byval value as ubyte)
declare sub format_set_diag_color(byval format as lxw_format ptr, byval color as lxw_color_t)
declare sub format_set_diag_border(byval format as lxw_format ptr, byval value as ubyte)
declare sub format_set_font_outline(byval format as lxw_format ptr)
declare sub format_set_font_shadow(byval format as lxw_format ptr)
declare sub format_set_font_family(byval format as lxw_format ptr, byval value as ubyte)
declare sub format_set_font_charset(byval format as lxw_format ptr, byval value as ubyte)
declare sub format_set_font_scheme(byval format as lxw_format ptr, byval font_scheme as const zstring ptr)
declare sub format_set_font_condense(byval format as lxw_format ptr)
declare sub format_set_font_extend(byval format as lxw_format ptr)
declare sub format_set_reading_order(byval format as lxw_format ptr, byval value as ubyte)
declare sub format_set_theme(byval format as lxw_format ptr, byval value as ubyte)
type lxw_chart_series as lxw_chart_series_

type lxw_chart_series_list
	stqh_first as lxw_chart_series ptr
	stqh_last as lxw_chart_series ptr ptr
end type

type lxw_series_data_point as lxw_series_data_point_

type lxw_series_data_points
	stqh_first as lxw_series_data_point ptr
	stqh_last as lxw_series_data_point ptr ptr
end type

const LXW_CHART_NUM_FORMAT_LEN = 128
const LXW_CHART_DEFAULT_GAP = 501

type lxw_chart_type as long
enum
	LXW_CHART_NONE = 0
	LXW_CHART_AREA
	LXW_CHART_AREA_STACKED
	LXW_CHART_AREA_STACKED_PERCENT
	LXW_CHART_BAR
	LXW_CHART_BAR_STACKED
	LXW_CHART_BAR_STACKED_PERCENT
	LXW_CHART_COLUMN
	LXW_CHART_COLUMN_STACKED
	LXW_CHART_COLUMN_STACKED_PERCENT
	LXW_CHART_DOUGHNUT
	LXW_CHART_LINE
	LXW_CHART_PIE
	LXW_CHART_SCATTER
	LXW_CHART_SCATTER_STRAIGHT
	LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS
	LXW_CHART_SCATTER_SMOOTH
	LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS
	LXW_CHART_RADAR
	LXW_CHART_RADAR_WITH_MARKERS
	LXW_CHART_RADAR_FILLED
end enum

type lxw_chart_legend_position as long
enum
	LXW_CHART_LEGEND_NONE = 0
	LXW_CHART_LEGEND_RIGHT
	LXW_CHART_LEGEND_LEFT
	LXW_CHART_LEGEND_TOP
	LXW_CHART_LEGEND_BOTTOM
	LXW_CHART_LEGEND_OVERLAY_RIGHT
	LXW_CHART_LEGEND_OVERLAY_LEFT
end enum

type lxw_chart_line_dash_type as long
enum
	LXW_CHART_LINE_DASH_SOLID = 0
	LXW_CHART_LINE_DASH_ROUND_DOT
	LXW_CHART_LINE_DASH_SQUARE_DOT
	LXW_CHART_LINE_DASH_DASH
	LXW_CHART_LINE_DASH_DASH_DOT
	LXW_CHART_LINE_DASH_LONG_DASH
	LXW_CHART_LINE_DASH_LONG_DASH_DOT
	LXW_CHART_LINE_DASH_LONG_DASH_DOT_DOT
	LXW_CHART_LINE_DASH_DOT
	LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT
	LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT_DOT
end enum

type lxw_chart_marker_type as long
enum
	LXW_CHART_MARKER_AUTOMATIC
	LXW_CHART_MARKER_NONE
	LXW_CHART_MARKER_SQUARE
	LXW_CHART_MARKER_DIAMOND
	LXW_CHART_MARKER_TRIANGLE
	LXW_CHART_MARKER_X
	LXW_CHART_MARKER_STAR
	LXW_CHART_MARKER_SHORT_DASH
	LXW_CHART_MARKER_LONG_DASH
	LXW_CHART_MARKER_CIRCLE
	LXW_CHART_MARKER_PLUS
end enum

type lxw_chart_pattern_type as long
enum
	LXW_CHART_PATTERN_NONE
	LXW_CHART_PATTERN_PERCENT_5
	LXW_CHART_PATTERN_PERCENT_10
	LXW_CHART_PATTERN_PERCENT_20
	LXW_CHART_PATTERN_PERCENT_25
	LXW_CHART_PATTERN_PERCENT_30
	LXW_CHART_PATTERN_PERCENT_40
	LXW_CHART_PATTERN_PERCENT_50
	LXW_CHART_PATTERN_PERCENT_60
	LXW_CHART_PATTERN_PERCENT_70
	LXW_CHART_PATTERN_PERCENT_75
	LXW_CHART_PATTERN_PERCENT_80
	LXW_CHART_PATTERN_PERCENT_90
	LXW_CHART_PATTERN_LIGHT_DOWNWARD_DIAGONAL
	LXW_CHART_PATTERN_LIGHT_UPWARD_DIAGONAL
	LXW_CHART_PATTERN_DARK_DOWNWARD_DIAGONAL
	LXW_CHART_PATTERN_DARK_UPWARD_DIAGONAL
	LXW_CHART_PATTERN_WIDE_DOWNWARD_DIAGONAL
	LXW_CHART_PATTERN_WIDE_UPWARD_DIAGONAL
	LXW_CHART_PATTERN_LIGHT_VERTICAL
	LXW_CHART_PATTERN_LIGHT_HORIZONTAL
	LXW_CHART_PATTERN_NARROW_VERTICAL
	LXW_CHART_PATTERN_NARROW_HORIZONTAL
	LXW_CHART_PATTERN_DARK_VERTICAL
	LXW_CHART_PATTERN_DARK_HORIZONTAL
	LXW_CHART_PATTERN_DASHED_DOWNWARD_DIAGONAL
	LXW_CHART_PATTERN_DASHED_UPWARD_DIAGONAL
	LXW_CHART_PATTERN_DASHED_HORIZONTAL
	LXW_CHART_PATTERN_DASHED_VERTICAL
	LXW_CHART_PATTERN_SMALL_CONFETTI
	LXW_CHART_PATTERN_LARGE_CONFETTI
	LXW_CHART_PATTERN_ZIGZAG
	LXW_CHART_PATTERN_WAVE
	LXW_CHART_PATTERN_DIAGONAL_BRICK
	LXW_CHART_PATTERN_HORIZONTAL_BRICK
	LXW_CHART_PATTERN_WEAVE
	LXW_CHART_PATTERN_PLAID
	LXW_CHART_PATTERN_DIVOT
	LXW_CHART_PATTERN_DOTTED_GRID
	LXW_CHART_PATTERN_DOTTED_DIAMOND
	LXW_CHART_PATTERN_SHINGLE
	LXW_CHART_PATTERN_TRELLIS
	LXW_CHART_PATTERN_SPHERE
	LXW_CHART_PATTERN_SMALL_GRID
	LXW_CHART_PATTERN_LARGE_GRID
	LXW_CHART_PATTERN_SMALL_CHECK
	LXW_CHART_PATTERN_LARGE_CHECK
	LXW_CHART_PATTERN_OUTLINED_DIAMOND
	LXW_CHART_PATTERN_SOLID_DIAMOND
end enum

type lxw_chart_label_position as long
enum
	LXW_CHART_LABEL_POSITION_DEFAULT
	LXW_CHART_LABEL_POSITION_CENTER
	LXW_CHART_LABEL_POSITION_RIGHT
	LXW_CHART_LABEL_POSITION_LEFT
	LXW_CHART_LABEL_POSITION_ABOVE
	LXW_CHART_LABEL_POSITION_BELOW
	LXW_CHART_LABEL_POSITION_INSIDE_BASE
	LXW_CHART_LABEL_POSITION_INSIDE_END
	LXW_CHART_LABEL_POSITION_OUTSIDE_END
	LXW_CHART_LABEL_POSITION_BEST_FIT
end enum

type lxw_chart_label_separator as long
enum
	LXW_CHART_LABEL_SEPARATOR_COMMA
	LXW_CHART_LABEL_SEPARATOR_SEMICOLON
	LXW_CHART_LABEL_SEPARATOR_PERIOD
	LXW_CHART_LABEL_SEPARATOR_NEWLINE
	LXW_CHART_LABEL_SEPARATOR_SPACE
end enum

type lxw_chart_subtype as long
enum
	LXW_CHART_SUBTYPE_NONE = 0
	LXW_CHART_SUBTYPE_STACKED
	LXW_CHART_SUBTYPE_STACKED_PERCENT
end enum

type lxw_chart_grouping as long
enum
	LXW_GROUPING_CLUSTERED
	LXW_GROUPING_STANDARD
	LXW_GROUPING_PERCENTSTACKED
	LXW_GROUPING_STACKED
end enum

type lxw_chart_axis_tick_position as long
enum
	LXW_CHART_AXIS_POSITION_DEFAULT
	LXW_CHART_AXIS_POSITION_ON_TICK
	LXW_CHART_AXIS_POSITION_BETWEEN
end enum

type lxw_chart_axis_label_position as long
enum
	LXW_CHART_AXIS_LABEL_POSITION_NEXT_TO
	LXW_CHART_AXIS_LABEL_POSITION_HIGH
	LXW_CHART_AXIS_LABEL_POSITION_LOW
	LXW_CHART_AXIS_LABEL_POSITION_NONE
end enum

type lxw_chart_axis_display_unit as long
enum
	LXW_CHART_AXIS_UNITS_NONE
	LXW_CHART_AXIS_UNITS_HUNDREDS
	LXW_CHART_AXIS_UNITS_THOUSANDS
	LXW_CHART_AXIS_UNITS_TEN_THOUSANDS
	LXW_CHART_AXIS_UNITS_HUNDRED_THOUSANDS
	LXW_CHART_AXIS_UNITS_MILLIONS
	LXW_CHART_AXIS_UNITS_TEN_MILLIONS
	LXW_CHART_AXIS_UNITS_HUNDRED_MILLIONS
	LXW_CHART_AXIS_UNITS_BILLIONS
	LXW_CHART_AXIS_UNITS_TRILLIONS
end enum

type lxw_chart_axis_tick_mark as long
enum
	LXW_CHART_AXIS_TICK_MARK_DEFAULT
	LXW_CHART_AXIS_TICK_MARK_NONE
	LXW_CHART_AXIS_TICK_MARK_INSIDE
	LXW_CHART_AXIS_TICK_MARK_OUTSIDE
	LXW_CHART_AXIS_TICK_MARK_CROSSING
end enum

type lxw_chart_tick_mark as lxw_chart_axis_tick_mark

type lxw_series_range
	formula as zstring ptr
	sheetname as zstring ptr
	first_row as lxw_row_t
	last_row as lxw_row_t
	first_col as lxw_col_t
	last_col as lxw_col_t
	ignore_cache as ubyte
	has_string_cache as ubyte
	num_data_points as ushort
	data_cache as lxw_series_data_points ptr
end type

type lxw_series_data_point_list_pointers
	stqe_next as lxw_series_data_point ptr
end type

type lxw_series_data_point_
	is_string as ubyte
	number as double
	string as zstring ptr
	no_data as ubyte
	list_pointers as lxw_series_data_point_list_pointers
end type

type lxw_chart_line
	color as lxw_color_t
	none as ubyte
	width as single
	dash_type as ubyte
	transparency as ubyte
	has_color as ubyte
end type

type lxw_chart_fill
	color as lxw_color_t
	none as ubyte
	transparency as ubyte
	has_color as ubyte
end type

type lxw_chart_pattern
	fg_color as lxw_color_t
	bg_color as lxw_color_t
	as ubyte type
	has_fg_color as ubyte
	has_bg_color as ubyte
end type

type lxw_chart_font
	name as zstring ptr
	size as ushort
	bold as ubyte
	italic as ubyte
	underline as ubyte
	rotation as long
	color as lxw_color_t
	pitch_family as ubyte
	charset as ubyte
	baseline as byte
	has_color as ubyte
end type

type lxw_chart_marker
	as ubyte type
	size as ubyte
	line as lxw_chart_line ptr
	fill as lxw_chart_fill ptr
	pattern as lxw_chart_pattern ptr
end type

type lxw_chart_legend
	font as lxw_chart_font ptr
	position as ubyte
end type

type lxw_chart_title
	name as zstring ptr
	row as lxw_row_t
	col as lxw_col_t
	font as lxw_chart_font ptr
	off as ubyte
	is_horizontal as ubyte
	ignore_cache as ubyte
	range as lxw_series_range ptr
	data_point as lxw_series_data_point
end type

type lxw_chart_point
	line as lxw_chart_line ptr
	fill as lxw_chart_fill ptr
	pattern as lxw_chart_pattern ptr
end type

type lxw_chart_blank as long
enum
	LXW_CHART_BLANKS_AS_GAP
	LXW_CHART_BLANKS_AS_ZERO
	LXW_CHART_BLANKS_AS_CONNECTED
end enum

type lxw_chart_position as long
enum
	LXW_CHART_AXIS_RIGHT
	LXW_CHART_AXIS_LEFT
	LXW_CHART_AXIS_TOP
	LXW_CHART_AXIS_BOTTOM
end enum

type lxw_chart_error_bar_type as long
enum
	LXW_CHART_ERROR_BAR_TYPE_STD_ERROR
	LXW_CHART_ERROR_BAR_TYPE_FIXED
	LXW_CHART_ERROR_BAR_TYPE_PERCENTAGE
	LXW_CHART_ERROR_BAR_TYPE_STD_DEV
end enum

type lxw_chart_error_bar_direction as long
enum
	LXW_CHART_ERROR_BAR_DIR_BOTH
	LXW_CHART_ERROR_BAR_DIR_PLUS
	LXW_CHART_ERROR_BAR_DIR_MINUS
end enum

type lxw_chart_error_bar_cap as long
enum
	LXW_CHART_ERROR_BAR_END_CAP
	LXW_CHART_ERROR_BAR_NO_CAP
end enum

type lxw_series_error_bars
	as ubyte type
	direction as ubyte
	endcap as ubyte
	has_value as ubyte
	is_set as ubyte
	is_x as ubyte
	chart_group as ubyte
	value as double
	line as lxw_chart_line ptr
end type

type lxw_chart_trendline_type as long
enum
	LXW_CHART_TRENDLINE_TYPE_LINEAR
	LXW_CHART_TRENDLINE_TYPE_LOG
	LXW_CHART_TRENDLINE_TYPE_POLY
	LXW_CHART_TRENDLINE_TYPE_POWER
	LXW_CHART_TRENDLINE_TYPE_EXP
	LXW_CHART_TRENDLINE_TYPE_AVERAGE
end enum

type lxw_chart_series_list_pointers
	stqe_next as lxw_chart_series ptr
end type

type lxw_chart_series_
	categories as lxw_series_range ptr
	values as lxw_series_range ptr
	title as lxw_chart_title
	line as lxw_chart_line ptr
	fill as lxw_chart_fill ptr
	pattern as lxw_chart_pattern ptr
	marker as lxw_chart_marker ptr
	points as lxw_chart_point ptr
	point_count as ushort
	smooth as ubyte
	invert_if_negative as ubyte
	has_labels as ubyte
	show_labels_value as ubyte
	show_labels_category as ubyte
	show_labels_name as ubyte
	show_labels_leader as ubyte
	show_labels_legend as ubyte
	show_labels_percent as ubyte
	label_position as ubyte
	label_separator as ubyte
	default_label_position as ubyte
	label_num_format as zstring ptr
	label_font as lxw_chart_font ptr
	x_error_bars as lxw_series_error_bars ptr
	y_error_bars as lxw_series_error_bars ptr
	has_trendline as ubyte
	has_trendline_forecast as ubyte
	has_trendline_equation as ubyte
	has_trendline_r_squared as ubyte
	has_trendline_intercept as ubyte
	trendline_type as ubyte
	trendline_value as ubyte
	trendline_forward as double
	trendline_backward as double
	trendline_value_type as ubyte
	trendline_name as zstring ptr
	trendline_line as lxw_chart_line ptr
	trendline_intercept as double
	list_pointers as lxw_chart_series_list_pointers
end type

type lxw_chart_gridline
	visible as ubyte
	line as lxw_chart_line ptr
end type

type lxw_chart_axis
	title as lxw_chart_title
	num_format as zstring ptr
	default_num_format as zstring ptr
	source_linked as ubyte
	major_tick_mark as ubyte
	minor_tick_mark as ubyte
	is_horizontal as ubyte
	major_gridlines as lxw_chart_gridline
	minor_gridlines as lxw_chart_gridline
	num_font as lxw_chart_font ptr
	line as lxw_chart_line ptr
	fill as lxw_chart_fill ptr
	pattern as lxw_chart_pattern ptr
	is_category as ubyte
	is_date as ubyte
	is_value as ubyte
	axis_position as ubyte
	position_axis as ubyte
	label_position as ubyte
	hidden as ubyte
	reverse as ubyte
	has_min as ubyte
	min as double
	has_max as ubyte
	max as double
	has_major_unit as ubyte
	major_unit as double
	has_minor_unit as ubyte
	minor_unit as double
	interval_unit as ushort
	interval_tick as ushort
	log_base as ushort
	display_units as ubyte
	display_units_visible as ubyte
	has_crossing as ubyte
	crossing_max as ubyte
	crossing as double
end type

type lxw_chart as lxw_chart_

type lxw_chart_ordered_list_pointers
	stqe_next as lxw_chart ptr
end type

type lxw_chart_list_pointers
	stqe_next as lxw_chart ptr
end type

type lxw_chart_
	file as FILE ptr
	as ubyte type
	subtype as ubyte
	series_index as ushort
	write_chart_type as sub(byval as lxw_chart ptr)
	write_plot_area as sub(byval as lxw_chart ptr)
	x_axis as lxw_chart_axis ptr
	y_axis as lxw_chart_axis ptr
	title as lxw_chart_title
	id as ulong
	axis_id_1 as ulong
	axis_id_2 as ulong
	axis_id_3 as ulong
	axis_id_4 as ulong
	in_use as ubyte
	chart_group as ubyte
	cat_has_num_fmt as ubyte
	has_horiz_cat_axis as ubyte
	has_horiz_val_axis as ubyte
	style_id as ubyte
	rotation as ushort
	hole_size as ushort
	no_title as ubyte
	has_overlap as ubyte
	overlap_y1 as byte
	overlap_y2 as byte
	gap_y1 as ushort
	gap_y2 as ushort
	grouping as ubyte
	default_cross_between as ubyte
	legend as lxw_chart_legend
	delete_series as short ptr
	delete_series_count as ushort
	default_marker as lxw_chart_marker ptr
	chartarea_line as lxw_chart_line ptr
	chartarea_fill as lxw_chart_fill ptr
	chartarea_pattern as lxw_chart_pattern ptr
	plotarea_line as lxw_chart_line ptr
	plotarea_fill as lxw_chart_fill ptr
	plotarea_pattern as lxw_chart_pattern ptr
	has_drop_lines as ubyte
	drop_lines_line as lxw_chart_line ptr
	has_high_low_lines as ubyte
	high_low_lines_line as lxw_chart_line ptr
	series_list as lxw_chart_series_list ptr
	has_table as ubyte
	has_table_vertical as ubyte
	has_table_horizontal as ubyte
	has_table_outline as ubyte
	has_table_legend_keys as ubyte
	table_font as lxw_chart_font ptr
	show_blanks_as as ubyte
	show_hidden_data as ubyte
	has_up_down_bars as ubyte
	up_bar_line as lxw_chart_line ptr
	down_bar_line as lxw_chart_line ptr
	up_bar_fill as lxw_chart_fill ptr
	down_bar_fill as lxw_chart_fill ptr
	default_label_position as ubyte
	ordered_list_pointers as lxw_chart_ordered_list_pointers
	list_pointers as lxw_chart_list_pointers
end type

declare function lxw_chart_new(byval type as ubyte) as lxw_chart ptr
declare sub lxw_chart_free(byval chart as lxw_chart ptr)
declare sub lxw_chart_assemble_xml_file(byval chart as lxw_chart ptr)
declare function chart_add_series(byval chart as lxw_chart ptr, byval categories as const zstring ptr, byval values as const zstring ptr) as lxw_chart_series ptr
declare sub chart_series_set_categories(byval series as lxw_chart_series ptr, byval sheetname as const zstring ptr, byval first_row as lxw_row_t, byval first_col as lxw_col_t, byval last_row as lxw_row_t, byval last_col as lxw_col_t)
declare sub chart_series_set_values(byval series as lxw_chart_series ptr, byval sheetname as const zstring ptr, byval first_row as lxw_row_t, byval first_col as lxw_col_t, byval last_row as lxw_row_t, byval last_col as lxw_col_t)
declare sub chart_series_set_name(byval series as lxw_chart_series ptr, byval name as const zstring ptr)
declare sub chart_series_set_name_range(byval series as lxw_chart_series ptr, byval sheetname as const zstring ptr, byval row as lxw_row_t, byval col as lxw_col_t)
declare sub chart_series_set_line(byval series as lxw_chart_series ptr, byval line as lxw_chart_line ptr)
declare sub chart_series_set_fill(byval series as lxw_chart_series ptr, byval fill as lxw_chart_fill ptr)
declare sub chart_series_set_invert_if_negative(byval series as lxw_chart_series ptr)
declare sub chart_series_set_pattern(byval series as lxw_chart_series ptr, byval pattern as lxw_chart_pattern ptr)
declare sub chart_series_set_marker_type(byval series as lxw_chart_series ptr, byval type as ubyte)
declare sub chart_series_set_marker_size(byval series as lxw_chart_series ptr, byval size as ubyte)
declare sub chart_series_set_marker_line(byval series as lxw_chart_series ptr, byval line as lxw_chart_line ptr)
declare sub chart_series_set_marker_fill(byval series as lxw_chart_series ptr, byval fill as lxw_chart_fill ptr)
declare sub chart_series_set_marker_pattern(byval series as lxw_chart_series ptr, byval pattern as lxw_chart_pattern ptr)
declare function chart_series_set_points(byval series as lxw_chart_series ptr, byval points as lxw_chart_point ptr ptr) as lxw_error
declare sub chart_series_set_smooth(byval series as lxw_chart_series ptr, byval smooth as ubyte)
declare sub chart_series_set_labels(byval series as lxw_chart_series ptr)
declare sub chart_series_set_labels_options(byval series as lxw_chart_series ptr, byval show_name as ubyte, byval show_category as ubyte, byval show_value as ubyte)
declare sub chart_series_set_labels_separator(byval series as lxw_chart_series ptr, byval separator as ubyte)
declare sub chart_series_set_labels_position(byval series as lxw_chart_series ptr, byval position as ubyte)
declare sub chart_series_set_labels_leader_line(byval series as lxw_chart_series ptr)
declare sub chart_series_set_labels_legend(byval series as lxw_chart_series ptr)
declare sub chart_series_set_labels_percentage(byval series as lxw_chart_series ptr)
declare sub chart_series_set_labels_num_format(byval series as lxw_chart_series ptr, byval num_format as const zstring ptr)
declare sub chart_series_set_labels_font(byval series as lxw_chart_series ptr, byval font as lxw_chart_font ptr)
declare sub chart_series_set_trendline(byval series as lxw_chart_series ptr, byval type as ubyte, byval value as ubyte)
declare sub chart_series_set_trendline_forecast(byval series as lxw_chart_series ptr, byval forward as double, byval backward as double)
declare sub chart_series_set_trendline_equation(byval series as lxw_chart_series ptr)
declare sub chart_series_set_trendline_r_squared(byval series as lxw_chart_series ptr)
declare sub chart_series_set_trendline_intercept(byval series as lxw_chart_series ptr, byval intercept as double)
declare sub chart_series_set_trendline_name(byval series as lxw_chart_series ptr, byval name as const zstring ptr)
declare sub chart_series_set_trendline_line(byval series as lxw_chart_series ptr, byval line as lxw_chart_line ptr)
declare sub chart_series_set_error_bars(byval error_bars as lxw_series_error_bars ptr, byval type as ubyte, byval value as double)
declare sub chart_series_set_error_bars_direction(byval error_bars as lxw_series_error_bars ptr, byval direction as ubyte)
declare sub chart_series_set_error_bars_endcap(byval error_bars as lxw_series_error_bars ptr, byval endcap as ubyte)
declare sub chart_series_set_error_bars_line(byval error_bars as lxw_series_error_bars ptr, byval line as lxw_chart_line ptr)
declare sub chart_axis_set_name(byval axis as lxw_chart_axis ptr, byval name as const zstring ptr)
declare sub chart_axis_set_name_range(byval axis as lxw_chart_axis ptr, byval sheetname as const zstring ptr, byval row as lxw_row_t, byval col as lxw_col_t)
declare sub chart_axis_set_name_font(byval axis as lxw_chart_axis ptr, byval font as lxw_chart_font ptr)
declare sub chart_axis_set_num_font(byval axis as lxw_chart_axis ptr, byval font as lxw_chart_font ptr)
declare sub chart_axis_set_num_format(byval axis as lxw_chart_axis ptr, byval num_format as const zstring ptr)
declare sub chart_axis_set_line(byval axis as lxw_chart_axis ptr, byval line as lxw_chart_line ptr)
declare sub chart_axis_set_fill(byval axis as lxw_chart_axis ptr, byval fill as lxw_chart_fill ptr)
declare sub chart_axis_set_pattern(byval axis as lxw_chart_axis ptr, byval pattern as lxw_chart_pattern ptr)
declare sub chart_axis_set_reverse(byval axis as lxw_chart_axis ptr)
declare sub chart_axis_set_crossing(byval axis as lxw_chart_axis ptr, byval value as double)
declare sub chart_axis_set_crossing_max(byval axis as lxw_chart_axis ptr)
declare sub chart_axis_off(byval axis as lxw_chart_axis ptr)
declare sub chart_axis_set_position(byval axis as lxw_chart_axis ptr, byval position as ubyte)
declare sub chart_axis_set_label_position(byval axis as lxw_chart_axis ptr, byval position as ubyte)
declare sub chart_axis_set_min(byval axis as lxw_chart_axis ptr, byval min as double)
declare sub chart_axis_set_max(byval axis as lxw_chart_axis ptr, byval max as double)
declare sub chart_axis_set_log_base(byval axis as lxw_chart_axis ptr, byval log_base as ushort)
declare sub chart_axis_set_major_tick_mark(byval axis as lxw_chart_axis ptr, byval type as ubyte)
declare sub chart_axis_set_minor_tick_mark(byval axis as lxw_chart_axis ptr, byval type as ubyte)
declare sub chart_axis_set_interval_unit(byval axis as lxw_chart_axis ptr, byval unit as ushort)
declare sub chart_axis_set_interval_tick(byval axis as lxw_chart_axis ptr, byval unit as ushort)
declare sub chart_axis_set_major_unit(byval axis as lxw_chart_axis ptr, byval unit as double)
declare sub chart_axis_set_minor_unit(byval axis as lxw_chart_axis ptr, byval unit as double)
declare sub chart_axis_set_display_units(byval axis as lxw_chart_axis ptr, byval units as ubyte)
declare sub chart_axis_set_display_units_visible(byval axis as lxw_chart_axis ptr, byval visible as ubyte)
declare sub chart_axis_major_gridlines_set_visible(byval axis as lxw_chart_axis ptr, byval visible as ubyte)
declare sub chart_axis_minor_gridlines_set_visible(byval axis as lxw_chart_axis ptr, byval visible as ubyte)
declare sub chart_axis_major_gridlines_set_line(byval axis as lxw_chart_axis ptr, byval line as lxw_chart_line ptr)
declare sub chart_axis_minor_gridlines_set_line(byval axis as lxw_chart_axis ptr, byval line as lxw_chart_line ptr)
declare sub chart_title_set_name(byval chart as lxw_chart ptr, byval name as const zstring ptr)
declare sub chart_title_set_name_range(byval chart as lxw_chart ptr, byval sheetname as const zstring ptr, byval row as lxw_row_t, byval col as lxw_col_t)
declare sub chart_title_set_name_font(byval chart as lxw_chart ptr, byval font as lxw_chart_font ptr)
declare sub chart_title_off(byval chart as lxw_chart ptr)
declare sub chart_legend_set_position(byval chart as lxw_chart ptr, byval position as ubyte)
declare sub chart_legend_set_font(byval chart as lxw_chart ptr, byval font as lxw_chart_font ptr)
declare function chart_legend_delete_series(byval chart as lxw_chart ptr, byval delete_series as short ptr) as lxw_error
declare sub chart_chartarea_set_line(byval chart as lxw_chart ptr, byval line as lxw_chart_line ptr)
declare sub chart_chartarea_set_fill(byval chart as lxw_chart ptr, byval fill as lxw_chart_fill ptr)
declare sub chart_chartarea_set_pattern(byval chart as lxw_chart ptr, byval pattern as lxw_chart_pattern ptr)
declare sub chart_plotarea_set_line(byval chart as lxw_chart ptr, byval line as lxw_chart_line ptr)
declare sub chart_plotarea_set_fill(byval chart as lxw_chart ptr, byval fill as lxw_chart_fill ptr)
declare sub chart_plotarea_set_pattern(byval chart as lxw_chart ptr, byval pattern as lxw_chart_pattern ptr)
declare sub chart_set_style(byval chart as lxw_chart ptr, byval style_id as ubyte)
declare sub chart_set_table(byval chart as lxw_chart ptr)
declare sub chart_set_table_grid(byval chart as lxw_chart ptr, byval horizontal as ubyte, byval vertical as ubyte, byval outline as ubyte, byval legend_keys as ubyte)
declare sub chart_set_table_font(byval chart as lxw_chart ptr, byval font as lxw_chart_font ptr)
declare sub chart_set_up_down_bars(byval chart as lxw_chart ptr)
declare sub chart_set_up_down_bars_format(byval chart as lxw_chart ptr, byval up_bar_line as lxw_chart_line ptr, byval up_bar_fill as lxw_chart_fill ptr, byval down_bar_line as lxw_chart_line ptr, byval down_bar_fill as lxw_chart_fill ptr)
declare sub chart_set_drop_lines(byval chart as lxw_chart ptr, byval line as lxw_chart_line ptr)
declare sub chart_set_high_low_lines(byval chart as lxw_chart ptr, byval line as lxw_chart_line ptr)
declare sub chart_set_series_overlap(byval chart as lxw_chart ptr, byval overlap as byte)
declare sub chart_set_series_gap(byval chart as lxw_chart ptr, byval gap as ushort)
declare sub chart_show_blanks_as(byval chart as lxw_chart ptr, byval option as ubyte)
declare sub chart_show_hidden_data(byval chart as lxw_chart ptr)
declare sub chart_set_rotation(byval chart as lxw_chart ptr, byval rotation as ushort)
declare sub chart_set_hole_size(byval chart as lxw_chart ptr, byval size as ubyte)
declare function lxw_chart_add_data_cache(byval range as lxw_series_range ptr, byval data as ubyte ptr, byval rows as ushort, byval cols as ubyte, byval col as ubyte) as lxw_error
#define __LXW_DRAWING_H__
type lxw_drawing_object as lxw_drawing_object_

type lxw_drawing_objects
	stqh_first as lxw_drawing_object ptr
	stqh_last as lxw_drawing_object ptr ptr
end type

type lxw_drawing_types as long
enum
	LXW_DRAWING_NONE = 0
	LXW_DRAWING_IMAGE
	LXW_DRAWING_CHART
	LXW_DRAWING_SHAPE
end enum

type lxw_anchor_types as long
enum
	LXW_ANCHOR_TYPE_NONE = 0
	LXW_ANCHOR_TYPE_IMAGE
	LXW_ANCHOR_TYPE_CHART
end enum

type lxw_anchor_edit_types as long
enum
	LXW_ANCHOR_EDIT_AS_NONE = 0
	LXW_ANCHOR_EDIT_AS_RELATIVE
	LXW_ANCHOR_EDIT_AS_ONE_CELL
	LXW_ANCHOR_EDIT_AS_ABSOLUTE
end enum

type image_types as long
enum
	LXW_IMAGE_UNKNOWN = 0
	LXW_IMAGE_PNG
	LXW_IMAGE_JPEG
	LXW_IMAGE_BMP
end enum

type lxw_drawing_coords
	col as ulong
	row as ulong
	col_offset as double
	row_offset as double
end type

type lxw_drawing_object_list_pointers
	stqe_next as lxw_drawing_object ptr
end type

type lxw_drawing_object_
	anchor_type as ubyte
	edit_as as ubyte
	from as lxw_drawing_coords
	to as lxw_drawing_coords
	col_absolute as ulong
	row_absolute as ulong
	width as ulong
	height as ulong
	shape as ubyte
	description as zstring ptr
	url as zstring ptr
	tip as zstring ptr
	list_pointers as lxw_drawing_object_list_pointers
end type

type lxw_drawing
	file as FILE ptr
	embedded as ubyte
	drawing_objects as lxw_drawing_objects ptr
end type

declare function lxw_drawing_new() as lxw_drawing ptr
declare sub lxw_drawing_free(byval drawing as lxw_drawing ptr)
declare sub lxw_drawing_assemble_xml_file(byval self as lxw_drawing ptr)
declare sub lxw_free_drawing_object(byval drawing_object as lxw_drawing_object ptr)
declare sub lxw_add_drawing_object(byval drawing as lxw_drawing ptr, byval drawing_object as lxw_drawing_object ptr)
#define __LXW_UTILITY_H__

#define CELL(cell) lxw_name_to_row(cell), lxw_name_to_col(cell)
#define COLS(cols) lxw_name_to_col(cols), lxw_name_to_col_2(cols)
#define RANGE(range) lxw_name_to_row(range), lxw_name_to_col(range), lxw_name_to_row_2(range), lxw_name_to_col_2(range)

declare function lxw_strerror(byval error_num as lxw_error) as zstring ptr
declare function lxw_quote_sheetname(byval str as const zstring ptr) as zstring ptr
declare sub lxw_col_to_name(byval col_name as zstring ptr, byval col_num as lxw_col_t, byval absolute as ubyte)
declare sub lxw_rowcol_to_cell(byval cell_name as zstring ptr, byval row as lxw_row_t, byval col as lxw_col_t)
declare sub lxw_rowcol_to_cell_abs(byval cell_name as zstring ptr, byval row as lxw_row_t, byval col as lxw_col_t, byval abs_row as ubyte, byval abs_col as ubyte)
declare sub lxw_rowcol_to_range(byval range as zstring ptr, byval first_row as lxw_row_t, byval first_col as lxw_col_t, byval last_row as lxw_row_t, byval last_col as lxw_col_t)
declare sub lxw_rowcol_to_range_abs(byval range as zstring ptr, byval first_row as lxw_row_t, byval first_col as lxw_col_t, byval last_row as lxw_row_t, byval last_col as lxw_col_t)
declare sub lxw_rowcol_to_formula_abs(byval formula as zstring ptr, byval sheetname as const zstring ptr, byval first_row as lxw_row_t, byval first_col as lxw_col_t, byval last_row as lxw_row_t, byval last_col as lxw_col_t)
declare function lxw_name_to_row(byval row_str as const zstring ptr) as ulong
declare function lxw_name_to_col(byval col_str as const zstring ptr) as ushort
declare function lxw_name_to_row_2(byval row_str as const zstring ptr) as ulong
declare function lxw_name_to_col_2(byval col_str as const zstring ptr) as ushort
declare function lxw_datetime_to_excel_date(byval datetime as lxw_datetime ptr, byval date_1904 as ubyte) as double
declare function lxw_strdup(byval str as const zstring ptr) as zstring ptr
declare function lxw_utf8_strlen(byval str as const zstring ptr) as uinteger
declare sub lxw_str_tolower(byval str as zstring ptr)
declare function lxw_tmpfile(byval tmpdir as zstring ptr) as FILE ptr

const LXW_ROW_MAX = 1048576
const LXW_COL_MAX = 16384
const LXW_COL_META_MAX = 128
const LXW_HEADER_FOOTER_MAX = 255
const LXW_MAX_NUMBER_URLS = 65530
const LXW_PANE_NAME_LENGTH = 12
const LXW_BREAKS_MAX = 1023
const LXW_DEF_COL_WIDTH = cdbl(8.43)
const LXW_DEF_ROW_HEIGHT = cdbl(15.0)

type lxw_gridlines as long
enum
	LXW_HIDE_ALL_GRIDLINES = 0
	LXW_SHOW_SCREEN_GRIDLINES
	LXW_SHOW_PRINT_GRIDLINES
	LXW_SHOW_ALL_GRIDLINES
end enum

type cell_types as long
enum
	NUMBER_CELL = 1
	STRING_CELL
	INLINE_STRING_CELL
	FORMULA_CELL
	ARRAY_FORMULA_CELL
	BLANK_CELL
	BOOLEAN_CELL
	HYPERLINK_URL
	HYPERLINK_INTERNAL
	HYPERLINK_EXTERNAL
end enum

type pane_types as long
enum
	NO_PANES = 0
	FREEZE_PANES
	SPLIT_PANES
	FREEZE_SPLIT_PANES
end enum

type lxw_cell as lxw_cell_

type lxw_table_cells
	rbh_root as lxw_cell ptr
end type

type lxw_row as lxw_row_

type lxw_table_rows
	rbh_root as lxw_row ptr
	cached_row as lxw_row ptr
	cached_row_num as lxw_row_t
end type

'' TODO: #define LXW_RB_GENERATE_ROW(name, type, field, cmp) RB_GENERATE_INSERT_COLOR(name, type, field, static) RB_GENERATE_REMOVE_COLOR(name, type, field, static) RB_GENERATE_INSERT(name, type, field, cmp, static) RB_GENERATE_REMOVE(name, type, field, static) RB_GENERATE_FIND(name, type, field, cmp, static) RB_GENERATE_NEXT(name, type, field, static) RB_GENERATE_MINMAX(name, type, field, static) struct lxw_rb_generate_row{int unused;}
'' TODO: #define LXW_RB_GENERATE_CELL(name, type, field, cmp) RB_GENERATE_INSERT_COLOR(name, type, field, static) RB_GENERATE_REMOVE_COLOR(name, type, field, static) RB_GENERATE_INSERT(name, type, field, cmp, static) RB_GENERATE_REMOVE(name, type, field, static) RB_GENERATE_FIND(name, type, field, cmp, static) RB_GENERATE_NEXT(name, type, field, static) RB_GENERATE_MINMAX(name, type, field, static) struct lxw_rb_generate_cell{int unused;}
type lxw_merged_range as lxw_merged_range_

type lxw_merged_ranges
	stqh_first as lxw_merged_range ptr
	stqh_last as lxw_merged_range ptr ptr
end type

type lxw_selection as lxw_selection_

type lxw_selections
	stqh_first as lxw_selection ptr
	stqh_last as lxw_selection ptr ptr
end type

type lxw_image_options as lxw_image_options_

type lxw_image_data
	stqh_first as lxw_image_options ptr
	stqh_last as lxw_image_options ptr ptr
end type

type lxw_chart_data
	stqh_first as lxw_image_options ptr
	stqh_last as lxw_image_options ptr ptr
end type

type lxw_row_col_options
	hidden as ubyte
	level as ubyte
	collapsed as ubyte
end type

type lxw_col_options
	firstcol as lxw_col_t
	lastcol as lxw_col_t
	width as double
	format as lxw_format ptr
	hidden as ubyte
	level as ubyte
	collapsed as ubyte
end type

type lxw_merged_range_list_pointers
	stqe_next as lxw_merged_range ptr
end type

type lxw_merged_range_
	first_row as lxw_row_t
	last_row as lxw_row_t
	first_col as lxw_col_t
	last_col as lxw_col_t
	list_pointers as lxw_merged_range_list_pointers
end type

type lxw_repeat_rows
	in_use as ubyte
	first_row as lxw_row_t
	last_row as lxw_row_t
end type

type lxw_repeat_cols
	in_use as ubyte
	first_col as lxw_col_t
	last_col as lxw_col_t
end type

type lxw_print_area
	in_use as ubyte
	first_row as lxw_row_t
	last_row as lxw_row_t
	first_col as lxw_col_t
	last_col as lxw_col_t
end type

type lxw_autofilter
	in_use as ubyte
	first_row as lxw_row_t
	last_row as lxw_row_t
	first_col as lxw_col_t
	last_col as lxw_col_t
end type

type lxw_panes
	as ubyte type
	first_row as lxw_row_t
	first_col as lxw_col_t
	top_row as lxw_row_t
	left_col as lxw_col_t
	x_split as double
	y_split as double
end type

type lxw_selection_list_pointers
	stqe_next as lxw_selection ptr
end type

type lxw_selection_
	pane as zstring * 12
	active_cell as zstring * sizeof("$XFWD$1048576") * 2
	sqref as zstring * sizeof("$XFWD$1048576") * 2
	list_pointers as lxw_selection_list_pointers
end type

type lxw_image_options_list_pointers
	stqe_next as lxw_image_options ptr
end type

type lxw_image_options_
	x_offset as long
	y_offset as long
	x_scale as double
	y_scale as double
	row as lxw_row_t
	col as lxw_col_t
	filename as zstring ptr
	url as zstring ptr
	tip as zstring ptr
	anchor as ubyte
	stream as FILE ptr
	image_type as ubyte
	width as double
	height as double
	short_name as zstring ptr
	extension as zstring ptr
	x_dpi as double
	y_dpi as double
	chart as lxw_chart ptr
	list_pointers as lxw_image_options_list_pointers
end type

type lxw_header_footer_options
	margin as double
end type

type lxw_protection
	no_select_locked_cells as ubyte
	no_select_unlocked_cells as ubyte
	format_cells as ubyte
	format_columns as ubyte
	format_rows as ubyte
	insert_columns as ubyte
	insert_rows as ubyte
	insert_hyperlinks as ubyte
	delete_columns as ubyte
	delete_rows as ubyte
	sort as ubyte
	autofilter as ubyte
	pivot_tables as ubyte
	scenarios as ubyte
	objects as ubyte
	no_sheet as ubyte
	content as ubyte
	is_configured as ubyte
	hash as zstring * 5
end type

type lxw_worksheet as lxw_worksheet_

type lxw_worksheet_list_pointers
	stqe_next as lxw_worksheet ptr
end type

type lxw_worksheet_
	file as FILE ptr
	optimize_tmpfile as FILE ptr
	table as lxw_table_rows ptr
	hyperlinks as lxw_table_rows ptr
	array as lxw_cell ptr ptr
	merged_ranges as lxw_merged_ranges ptr
	selections as lxw_selections ptr
	image_data as lxw_image_data ptr
	chart_data as lxw_chart_data ptr
	dim_rowmin as lxw_row_t
	dim_rowmax as lxw_row_t
	dim_colmin as lxw_col_t
	dim_colmax as lxw_col_t
	sst as lxw_sst ptr
	name as zstring ptr
	quoted_name as zstring ptr
	tmpdir as zstring ptr
	index as ulong
	active as ubyte
	selected as ubyte
	hidden as ubyte
	active_sheet as ushort ptr
	first_sheet as ushort ptr
	col_options as lxw_col_options ptr ptr
	col_options_max as ushort
	col_sizes as double ptr
	col_sizes_max as ushort
	col_formats as lxw_format ptr ptr
	col_formats_max as ushort
	col_size_changed as ubyte
	row_size_changed as ubyte
	optimize as ubyte
	optimize_row as lxw_row ptr
	fit_height as ushort
	fit_width as ushort
	horizontal_dpi as ushort
	hlink_count as ushort
	page_start as ushort
	print_scale as ushort
	rel_count as ushort
	vertical_dpi as ushort
	zoom as ushort
	filter_on as ubyte
	fit_page as ubyte
	hcenter as ubyte
	orientation as ubyte
	outline_changed as ubyte
	outline_on as ubyte
	page_order as ubyte
	page_setup_changed as ubyte
	page_view as ubyte
	paper_size as ubyte
	print_gridlines as ubyte
	print_headers as ubyte
	print_options_changed as ubyte
	right_to_left as ubyte
	screen_gridlines as ubyte
	show_zeros as ubyte
	vba_codename as ubyte
	vcenter as ubyte
	zoom_scale_normal as ubyte
	tab_color as lxw_color_t
	margin_left as double
	margin_right as double
	margin_top as double
	margin_bottom as double
	margin_header as double
	margin_footer as double
	default_row_height as double
	default_row_pixels as ulong
	default_col_pixels as ulong
	default_row_zeroed as ubyte
	default_row_set as ubyte
	header_footer_changed as ubyte
	header as zstring * 255
	footer as zstring * 255
	repeat_rows as lxw_repeat_rows
	repeat_cols as lxw_repeat_cols
	print_area as lxw_print_area
	autofilter as lxw_autofilter
	merged_range_count as ushort
	hbreaks as lxw_row_t ptr
	vbreaks as lxw_col_t ptr
	hbreaks_count as ushort
	vbreaks_count as ushort
	' ~ external_hyperlinks as lxw_rel_tuples ptr
	' ~ external_drawing_links as lxw_rel_tuples ptr
	' ~ drawing_links as lxw_rel_tuples ptr
	panes as lxw_panes
	protection as lxw_protection
	drawing as lxw_drawing ptr
	list_pointers as lxw_worksheet_list_pointers
end type

type lxw_worksheet_init_data
	index as ulong
	hidden as ubyte
	optimize as ubyte
	active_sheet as ushort ptr
	first_sheet as ushort ptr
	sst as lxw_sst ptr
	name as zstring ptr
	quoted_name as zstring ptr
	tmpdir as zstring ptr
end type

type lxw_row_tree_pointers
	rbe_left as lxw_row ptr
	rbe_right as lxw_row ptr
	rbe_parent as lxw_row ptr
	rbe_color as long
end type

type lxw_row_
	row_num as lxw_row_t
	height as double
	format as lxw_format ptr
	hidden as ubyte
	level as ubyte
	collapsed as ubyte
	row_changed as ubyte
	data_changed as ubyte
	height_changed as ubyte
	cells as lxw_table_cells ptr
	tree_pointers as lxw_row_tree_pointers
end type

union lxw_cell_u
	number as double
	string_id as long
	string as zstring ptr
end union

type lxw_cell_tree_pointers
	rbe_left as lxw_cell ptr
	rbe_right as lxw_cell ptr
	rbe_parent as lxw_cell ptr
	rbe_color as long
end type

type lxw_cell_
	row_num as lxw_row_t
	col_num as lxw_col_t
	as cell_types type
	format as lxw_format ptr
	u as lxw_cell_u
	formula_result as double
	user_data1 as zstring ptr
	user_data2 as zstring ptr
	sst_string as zstring ptr
	tree_pointers as lxw_cell_tree_pointers
end type

declare function worksheet_write_number(byval worksheet as lxw_worksheet ptr, byval row as lxw_row_t, byval col as lxw_col_t, byval number as double, byval format as lxw_format ptr) as lxw_error
declare function worksheet_write_string(byval worksheet as lxw_worksheet ptr, byval row as lxw_row_t, byval col as lxw_col_t, byval string as const zstring ptr, byval format as lxw_format ptr) as lxw_error
declare function worksheet_write_formula(byval worksheet as lxw_worksheet ptr, byval row as lxw_row_t, byval col as lxw_col_t, byval formula as const zstring ptr, byval format as lxw_format ptr) as lxw_error
declare function worksheet_write_array_formula(byval worksheet as lxw_worksheet ptr, byval first_row as lxw_row_t, byval first_col as lxw_col_t, byval last_row as lxw_row_t, byval last_col as lxw_col_t, byval formula as const zstring ptr, byval format as lxw_format ptr) as lxw_error
declare function worksheet_write_array_formula_num(byval worksheet as lxw_worksheet ptr, byval first_row as lxw_row_t, byval first_col as lxw_col_t, byval last_row as lxw_row_t, byval last_col as lxw_col_t, byval formula as const zstring ptr, byval format as lxw_format ptr, byval result as double) as lxw_error
declare function worksheet_write_datetime(byval worksheet as lxw_worksheet ptr, byval row as lxw_row_t, byval col as lxw_col_t, byval datetime as lxw_datetime ptr, byval format as lxw_format ptr) as lxw_error
declare function worksheet_write_url_opt(byval worksheet as lxw_worksheet ptr, byval row_num as lxw_row_t, byval col_num as lxw_col_t, byval url as const zstring ptr, byval format as lxw_format ptr, byval string as const zstring ptr, byval tooltip as const zstring ptr) as lxw_error
declare function worksheet_write_url(byval worksheet as lxw_worksheet ptr, byval row as lxw_row_t, byval col as lxw_col_t, byval url as const zstring ptr, byval format as lxw_format ptr) as lxw_error
declare function worksheet_write_boolean(byval worksheet as lxw_worksheet ptr, byval row as lxw_row_t, byval col as lxw_col_t, byval value as long, byval format as lxw_format ptr) as lxw_error
declare function worksheet_write_blank(byval worksheet as lxw_worksheet ptr, byval row as lxw_row_t, byval col as lxw_col_t, byval format as lxw_format ptr) as lxw_error
declare function worksheet_write_formula_num(byval worksheet as lxw_worksheet ptr, byval row as lxw_row_t, byval col as lxw_col_t, byval formula as const zstring ptr, byval format as lxw_format ptr, byval result as double) as lxw_error
declare function worksheet_set_row(byval worksheet as lxw_worksheet ptr, byval row as lxw_row_t, byval height as double, byval format as lxw_format ptr) as lxw_error
declare function worksheet_set_row_opt(byval worksheet as lxw_worksheet ptr, byval row as lxw_row_t, byval height as double, byval format as lxw_format ptr, byval options as lxw_row_col_options ptr) as lxw_error
declare function worksheet_set_column(byval worksheet as lxw_worksheet ptr, byval first_col as lxw_col_t, byval last_col as lxw_col_t, byval width as double, byval format as lxw_format ptr) as lxw_error
declare function worksheet_set_column_opt(byval worksheet as lxw_worksheet ptr, byval first_col as lxw_col_t, byval last_col as lxw_col_t, byval width as double, byval format as lxw_format ptr, byval options as lxw_row_col_options ptr) as lxw_error
declare function worksheet_insert_image(byval worksheet as lxw_worksheet ptr, byval row as lxw_row_t, byval col as lxw_col_t, byval filename as const zstring ptr) as lxw_error
declare function worksheet_insert_image_opt(byval worksheet as lxw_worksheet ptr, byval row as lxw_row_t, byval col as lxw_col_t, byval filename as const zstring ptr, byval options as lxw_image_options ptr) as lxw_error
declare function worksheet_insert_chart(byval worksheet as lxw_worksheet ptr, byval row as lxw_row_t, byval col as lxw_col_t, byval chart as lxw_chart ptr) as lxw_error
declare function worksheet_insert_chart_opt(byval worksheet as lxw_worksheet ptr, byval row as lxw_row_t, byval col as lxw_col_t, byval chart as lxw_chart ptr, byval user_options as lxw_image_options ptr) as lxw_error
declare function worksheet_merge_range(byval worksheet as lxw_worksheet ptr, byval first_row as lxw_row_t, byval first_col as lxw_col_t, byval last_row as lxw_row_t, byval last_col as lxw_col_t, byval string as const zstring ptr, byval format as lxw_format ptr) as lxw_error
declare function worksheet_autofilter(byval worksheet as lxw_worksheet ptr, byval first_row as lxw_row_t, byval first_col as lxw_col_t, byval last_row as lxw_row_t, byval last_col as lxw_col_t) as lxw_error
declare sub worksheet_activate(byval worksheet as lxw_worksheet ptr)
declare sub worksheet_select(byval worksheet as lxw_worksheet ptr)
declare sub worksheet_hide(byval worksheet as lxw_worksheet ptr)
declare sub worksheet_set_first_sheet(byval worksheet as lxw_worksheet ptr)
declare sub worksheet_freeze_panes(byval worksheet as lxw_worksheet ptr, byval row as lxw_row_t, byval col as lxw_col_t)
declare sub worksheet_split_panes(byval worksheet as lxw_worksheet ptr, byval vertical as double, byval horizontal as double)
declare sub worksheet_freeze_panes_opt(byval worksheet as lxw_worksheet ptr, byval first_row as lxw_row_t, byval first_col as lxw_col_t, byval top_row as lxw_row_t, byval left_col as lxw_col_t, byval type as ubyte)
declare sub worksheet_split_panes_opt(byval worksheet as lxw_worksheet ptr, byval vertical as double, byval horizontal as double, byval top_row as lxw_row_t, byval left_col as lxw_col_t)
declare sub worksheet_set_selection(byval worksheet as lxw_worksheet ptr, byval first_row as lxw_row_t, byval first_col as lxw_col_t, byval last_row as lxw_row_t, byval last_col as lxw_col_t)
declare sub worksheet_set_landscape(byval worksheet as lxw_worksheet ptr)
declare sub worksheet_set_portrait(byval worksheet as lxw_worksheet ptr)
declare sub worksheet_set_page_view(byval worksheet as lxw_worksheet ptr)
declare sub worksheet_set_paper(byval worksheet as lxw_worksheet ptr, byval paper_type as ubyte)
declare sub worksheet_set_margins(byval worksheet as lxw_worksheet ptr, byval left as double, byval right as double, byval top as double, byval bottom as double)
declare function worksheet_set_header(byval worksheet as lxw_worksheet ptr, byval string as const zstring ptr) as lxw_error
declare function worksheet_set_footer(byval worksheet as lxw_worksheet ptr, byval string as const zstring ptr) as lxw_error
declare function worksheet_set_header_opt(byval worksheet as lxw_worksheet ptr, byval string as const zstring ptr, byval options as lxw_header_footer_options ptr) as lxw_error
declare function worksheet_set_footer_opt(byval worksheet as lxw_worksheet ptr, byval string as const zstring ptr, byval options as lxw_header_footer_options ptr) as lxw_error
declare function worksheet_set_h_pagebreaks(byval worksheet as lxw_worksheet ptr, byval breaks as lxw_row_t ptr) as lxw_error
declare function worksheet_set_v_pagebreaks(byval worksheet as lxw_worksheet ptr, byval breaks as lxw_col_t ptr) as lxw_error
declare sub worksheet_print_across(byval worksheet as lxw_worksheet ptr)
declare sub worksheet_set_zoom(byval worksheet as lxw_worksheet ptr, byval scale as ushort)
declare sub worksheet_gridlines(byval worksheet as lxw_worksheet ptr, byval option as ubyte)
declare sub worksheet_center_horizontally(byval worksheet as lxw_worksheet ptr)
declare sub worksheet_center_vertically(byval worksheet as lxw_worksheet ptr)
declare sub worksheet_print_row_col_headers(byval worksheet as lxw_worksheet ptr)
declare function worksheet_repeat_rows(byval worksheet as lxw_worksheet ptr, byval first_row as lxw_row_t, byval last_row as lxw_row_t) as lxw_error
declare function worksheet_repeat_columns(byval worksheet as lxw_worksheet ptr, byval first_col as lxw_col_t, byval last_col as lxw_col_t) as lxw_error
declare function worksheet_print_area(byval worksheet as lxw_worksheet ptr, byval first_row as lxw_row_t, byval first_col as lxw_col_t, byval last_row as lxw_row_t, byval last_col as lxw_col_t) as lxw_error
declare sub worksheet_fit_to_pages(byval worksheet as lxw_worksheet ptr, byval width as ushort, byval height as ushort)
declare sub worksheet_set_start_page(byval worksheet as lxw_worksheet ptr, byval start_page as ushort)
declare sub worksheet_set_print_scale(byval worksheet as lxw_worksheet ptr, byval scale as ushort)
declare sub worksheet_right_to_left(byval worksheet as lxw_worksheet ptr)
declare sub worksheet_hide_zero(byval worksheet as lxw_worksheet ptr)
declare sub worksheet_set_tab_color(byval worksheet as lxw_worksheet ptr, byval color as lxw_color_t)
declare sub worksheet_protect(byval worksheet as lxw_worksheet ptr, byval password as const zstring ptr, byval options as lxw_protection ptr)
declare sub worksheet_set_default_row(byval worksheet as lxw_worksheet ptr, byval height as double, byval hide_unused_rows as ubyte)
declare function lxw_worksheet_new(byval init_data as lxw_worksheet_init_data ptr) as lxw_worksheet ptr
declare sub lxw_worksheet_free(byval worksheet as lxw_worksheet ptr)
declare sub lxw_worksheet_assemble_xml_file(byval worksheet as lxw_worksheet ptr)
declare sub lxw_worksheet_write_single_row(byval worksheet as lxw_worksheet ptr)
declare sub lxw_worksheet_prepare_image(byval worksheet as lxw_worksheet ptr, byval image_ref_id as ushort, byval drawing_id as ushort, byval image_data as lxw_image_options ptr)
declare sub lxw_worksheet_prepare_chart(byval worksheet as lxw_worksheet ptr, byval chart_ref_id as ushort, byval drawing_id as ushort, byval image_data as lxw_image_options ptr)
declare function lxw_worksheet_find_row(byval worksheet as lxw_worksheet ptr, byval row_num as lxw_row_t) as lxw_row ptr
declare function lxw_worksheet_find_cell(byval row as lxw_row ptr, byval col_num as lxw_col_t) as lxw_cell ptr
const LXW_DEFINED_NAME_LENGTH = 128
type lxw_worksheet_name as lxw_worksheet_name_

type lxw_worksheet_names
	rbh_root as lxw_worksheet_name ptr
end type

type lxw_worksheets
	stqh_first as lxw_worksheet ptr
	stqh_last as lxw_worksheet ptr ptr
end type

type lxw_charts
	stqh_first as lxw_chart ptr
	stqh_last as lxw_chart ptr ptr
end type

type lxw_defined_name as lxw_defined_name_

type lxw_defined_names
	tqh_first as lxw_defined_name ptr
	tqh_last as lxw_defined_name ptr ptr
end type

type lxw_worksheet_name_tree_pointers
	rbe_left as lxw_worksheet_name ptr
	rbe_right as lxw_worksheet_name ptr
	rbe_parent as lxw_worksheet_name ptr
	rbe_color as long
end type

type lxw_worksheet_name_
	name as const zstring ptr
	worksheet as lxw_worksheet ptr
	tree_pointers as lxw_worksheet_name_tree_pointers
end type

'' TODO: #define LXW_RB_GENERATE_NAMES(name, type, field, cmp) RB_GENERATE_INSERT_COLOR(name, type, field, static) RB_GENERATE_REMOVE_COLOR(name, type, field, static) RB_GENERATE_INSERT(name, type, field, cmp, static) RB_GENERATE_REMOVE(name, type, field, static) RB_GENERATE_FIND(name, type, field, cmp, static) RB_GENERATE_NEXT(name, type, field, static) RB_GENERATE_MINMAX(name, type, field, static) struct lxw_rb_generate_names{int unused;}
#define LXW_FOREACH_WORKSHEET(worksheet, workbook) STAILQ_FOREACH((worksheet), (workbook)->worksheets, list_pointers)

type lxw_defined_name_list_pointers
	tqe_next as lxw_defined_name ptr
	tqe_prev as lxw_defined_name ptr ptr
end type

type lxw_defined_name_
	index as short
	hidden as ubyte
	name as zstring * 128
	app_name as zstring * 128
	formula as zstring * 128
	normalised_name as zstring * 128
	normalised_sheetname as zstring * 128
	list_pointers as lxw_defined_name_list_pointers
end type

type lxw_doc_properties
	title as zstring ptr
	subject as zstring ptr
	author as zstring ptr
	manager as zstring ptr
	company as zstring ptr
	category as zstring ptr
	keywords as zstring ptr
	comments as zstring ptr
	status as zstring ptr
	hyperlink_base as zstring ptr
	created as time_t
end type

type lxw_workbook_options
	constant_memory as ubyte
	tmpdir as zstring ptr
end type

type lxw_workbook
	file as FILE ptr
	worksheets as lxw_worksheets ptr
	worksheet_names as lxw_worksheet_names ptr
	charts as lxw_charts ptr
	ordered_charts as lxw_charts ptr
	formats as lxw_formats ptr
	defined_names as lxw_defined_names ptr
	sst as lxw_sst ptr
	properties as lxw_doc_properties ptr
	custom_properties as lxw_custom_properties ptr
	filename as zstring ptr
	options as lxw_workbook_options
	num_sheets as ushort
	first_sheet as ushort
	active_sheet as ushort
	num_xf_formats as ushort
	num_format_count as ushort
	drawing_count as ushort
	font_count as ushort
	border_count as ushort
	fill_count as ushort
	optimize as ubyte
	has_png as ubyte
	has_jpeg as ubyte
	has_bmp as ubyte
	used_xf_formats as lxw_hash_table ptr
end type

declare function workbook_new(byval filename as const zstring ptr) as lxw_workbook ptr
declare function workbook_new_opt(byval filename as const zstring ptr, byval options as lxw_workbook_options ptr) as lxw_workbook ptr
declare function new_workbook(byval filename as const zstring ptr) as lxw_workbook ptr
declare function new_workbook_opt(byval filename as const zstring ptr, byval options as lxw_workbook_options ptr) as lxw_workbook ptr
declare function workbook_add_worksheet(byval workbook as lxw_workbook ptr, byval sheetname as const zstring ptr) as lxw_worksheet ptr
declare function workbook_add_format(byval workbook as lxw_workbook ptr) as lxw_format ptr
declare function workbook_add_chart(byval workbook as lxw_workbook ptr, byval chart_type as ubyte) as lxw_chart ptr
declare function workbook_close(byval workbook as lxw_workbook ptr) as lxw_error
declare function workbook_set_properties(byval workbook as lxw_workbook ptr, byval properties as lxw_doc_properties ptr) as lxw_error
declare function workbook_set_custom_property_string(byval workbook as lxw_workbook ptr, byval name as const zstring ptr, byval value as const zstring ptr) as lxw_error
declare function workbook_set_custom_property_number(byval workbook as lxw_workbook ptr, byval name as const zstring ptr, byval value as double) as lxw_error
declare function workbook_set_custom_property_integer(byval workbook as lxw_workbook ptr, byval name as const zstring ptr, byval value as long) as lxw_error
declare function workbook_set_custom_property_boolean(byval workbook as lxw_workbook ptr, byval name as const zstring ptr, byval value as ubyte) as lxw_error
declare function workbook_set_custom_property_datetime(byval workbook as lxw_workbook ptr, byval name as const zstring ptr, byval datetime as lxw_datetime ptr) as lxw_error
declare function workbook_define_name(byval workbook as lxw_workbook ptr, byval name as const zstring ptr, byval formula as const zstring ptr) as lxw_error
declare function workbook_get_worksheet_by_name(byval workbook as lxw_workbook ptr, byval name as const zstring ptr) as lxw_worksheet ptr
declare function workbook_validate_worksheet_name(byval workbook as lxw_workbook ptr, byval sheetname as const zstring ptr) as lxw_error
declare sub lxw_workbook_free(byval workbook as lxw_workbook ptr)
declare sub lxw_workbook_assemble_xml_file(byval workbook as lxw_workbook ptr)
declare sub lxw_workbook_set_default_xf_indices(byval workbook as lxw_workbook ptr)
#define LXW_VERSION "0.7.0"

end extern
