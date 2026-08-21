"""B3 GUI regression: Base-only insertion, public cache-only only-diff."""
from __future__ import annotations
import argparse
from contextlib import contextmanager
import hashlib
import json
import tempfile
import time
from pathlib import Path
from openpyxl import Workbook
import sow_merge_tool as sm

_CASE = "base-insert-only-diff"
_SHEET = "S1"
_ROW_COUNT = 2200
# Declaration/type rows are immutable comparison pairs, not data rows.
_SCHEMA_ROW_COUNT = 2
_INSERT_ID = 10

def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()

def _path_snapshot(path: Path) -> tuple[bool, bytes | None]:
    return (True, path.read_bytes()) if path.exists() else (False, None)

def _canon(value):
    if value is None or isinstance(value, (bool, int, float, str, bytes)): return value
    if isinstance(value, dict): return tuple(sorted(((_canon(k), _canon(v)) for k, v in value.items()), key=repr))
    if isinstance(value, (tuple, list, set, frozenset)):
        items = tuple(_canon(item) for item in value)
        return tuple(sorted(items, key=repr)) if isinstance(value, (set, frozenset)) else items
    return ("opaque", type(value).__name__, id(value))

def _digest(value) -> str: return hashlib.sha256(repr(_canon(value)).encode("utf-8")).hexdigest()
def _fact(value) -> dict:
    try: length = len(value)
    except Exception: length = None
    preview = repr(_canon(value))
    return {"type": type(value).__name__, "len": length, "digest": _digest(value), "preview": preview[:160]}

def _diff(before, after, path="hard") -> list[dict]:
    if before == after: return []
    if isinstance(before, dict) and isinstance(after, dict):
        result=[]
        for key in sorted(set(before)|set(after), key=repr):
            child=f"{path}[{key!r}]"
            if key not in before or key not in after: result.append({"path": child, "before": _fact(before.get(key)), "after": _fact(after.get(key))})
            else: result.extend(_diff(before[key], after[key], child))
        return result
    if isinstance(before, (tuple,list)) and isinstance(after,(tuple,list)):
        result=[]
        if len(before)!=len(after): result.append({"path":f"{path}.len","before":_fact(before),"after":_fact(after)})
        for index,(left,right) in enumerate(zip(before,after)): result.extend(_diff(left,right,f"{path}[{index}]"))
        return result
    return [{"path":path,"before":_fact(before),"after":_fact(after)}]

def _assert_same(before, after, action: str) -> None:
    changes=_diff(before,after)
    assert not changes, f"{action}: forbidden hard mutation {json.dumps(changes,ensure_ascii=False,sort_keys=True)}"

def _model_projection_fact(view)->dict:
    cache=view.column_comparison_cache; projection=view.column_projection; model=cache.model
    assert projection.model is model
    key=model.cache_key
    slots=tuple((int(slot.logical_idx),slot.mine_col,slot.base_col,slot.theirs_col,str(slot.state),str(slot.origin_side or ""),float(getattr(slot.confidence,"score",0.0)),str(getattr(slot.confidence,"reason","") or ""),tuple(getattr(slot.confidence,"cause_codes",()) or ())) for slot in model.slots)
    semantics={"key":(key.sheet_name,key.row_model_version,key.column_model_version,key.mine_edit_version,key.base_edit_version,key.theirs_edit_version),"slots":slots,"structural":tuple(sorted(cache.structural_diff_cols)),"unresolved":tuple(sorted(cache.unresolved_cols)),"projection_blocks":tuple(projection.block_ordinal_by_slot),"maps":tuple((name,tuple(getattr(model,name).entries)) for name in ("mine_physical_to_logical","base_physical_to_logical","theirs_physical_to_logical","mine_logical_to_physical","base_logical_to_physical","theirs_logical_to_physical"))}
    return {"identity":(id(cache),id(model),id(projection)),"bounds":(view.max_row,view.max_col,view.col_max_a,view.col_max_b,view.col_max_base),"flags":(view._align_rows_enabled,view._sheet_structural_diff,view._only_diff_source_version),"semantics":_fact(semantics)}

def _hard_snapshot(app,view,paths:dict[str,Path])->dict:
    overlays=app.sheet_operation_overlays or {}; overlay=overlays.get(view.sheet)
    row_maps={"pairs":tuple(tuple(pair) for pair in view.row_pairs),"a_to_pair":dict(view.row_a_to_pair_idx),"b_to_pair":dict(view.row_b_to_pair_idx),"mine_base":dict(view.mine_to_base_row),"theirs_base":dict(view.theirs_to_base_row),"base_override":dict(view.pair_base_row_override),"missing_base":dict(view._missing_base_row_map)}
    return {"input_sha":tuple(sorted((name,_sha256(path)) for name,path in paths.items())),"manual":_fact({name:getattr(app,name) for name in ("manual_a_cell_ops","manual_b_cell_ops","manual_a_formula_cache_ops","manual_b_formula_cache_ops","manual_a_row_ops","manual_b_row_ops","manual_a_column_ops","manual_b_column_ops","manual_sheet_ops","auto_sheet_ops")}),"undo_redo":_fact((app.undo_stack,app.redo_stack)),"modified_touched":_fact((app.modified_a,app.modified_b,app.modified_sheets_a,app.modified_sheets_b,app.user_touched_conflicts,view.touched_rows)),"overlay":_fact({name:(getattr(item,"topology_generation",None),getattr(item,"mutation_generation",None),getattr(item,"cells",None)) for name,item in overlays.items()}),"row_maps":_fact(row_maps),"prepared":{"raw_a":_fact(dict(view.pair_raw_parts_a)),"raw_b":_fact(dict(view.pair_raw_parts_b)),"raw_base":_fact(dict(view.pair_raw_parts_base)),"ab_diffs":_fact(dict(view.pair_diff_cols)),"base_diffs":_fact(dict(view.pair_base_diff_cols)),"exactness":(view._prepared_complete,view._data_ready,view._row_model_exact,view._pair_diff_full_exact,view._base_diff_full_exact)},"model_projection":_model_projection_fact(view),"generations":(app._sheet_compute_generation.get(view.sheet),view._row_model_version,view._column_model_version,view._column_projection_generation,view._virtual_column_window_generation,getattr(overlay,"topology_generation",None),getattr(overlay,"mutation_generation",None)),"edit_handles":tuple((name,id(getattr(app,name,None)),type(getattr(app,name,None)).__name__,getattr(getattr(app,name,None),"read_only",None)) for name in ("_wb_a_edit","_wb_b_edit","_wb_base_edit"))}

def _make_book(path:Path,include_insert:bool)->None:
    workbook=Workbook(); sheet=workbook.active; sheet.title=_SHEET
    sheet.append(["id@id","value"]); sheet.append(["int32","string"])
    for row_id in range(1,_ROW_COUNT+1):
        if include_insert or row_id!=_INSERT_ID: sheet.append([row_id,f"value-{row_id}"])
    workbook.save(path); workbook.close()

def _pump(app)->None: app.root.update_idletasks(); app.root.update()
def _wait(app,predicate,deadline:float,stage:str)->None:
    while time.monotonic()<deadline:
        _pump(app)
        if predicate(): return
        view=app.sheet_views.get(_SHEET)
        if view is not None and view._derive_lifecycle_state() in {"FAILED","UNRESOLVED","CANCELED","CLOSING"}: raise AssertionError(f"{stage}: terminal={view._derive_lifecycle_state()} error={view._lifecycle_error!r} entry={app._sheet_exact_entry(_SHEET)!r}")
        time.sleep(.005)
    view=app.sheet_views.get(_SHEET); facts=None if view is None else {"ready":view._data_ready,"prepared":view._prepared_complete,"only_diff":view.only_diff_var.get(),"pending":view._mode_switch_pending}
    raise AssertionError(f"timeout {stage}: entry={app._sheet_exact_entry(_SHEET)!r} selected={app.selected_sheet!r} view={facts!r}")

def _new_app(mine:Path,theirs:Path,base:Path,merged:Path):
    original=sm.SowMergeApp._schedule_formula_cache_prompt; sm.SowMergeApp._schedule_formula_cache_prompt=lambda _self:None
    try: return sm.SowMergeApp(str(mine),str(theirs),merge_mode=True,merged_path=str(merged),base_path=str(base),initial_sheet=_SHEET),original
    except Exception: sm.SowMergeApp._schedule_formula_cache_prompt=original; raise

def _cancel_debounce(app)->None:
    if app is not None:
        for view in tuple(app.sheet_views.values()):
            if view is not None and getattr(view,"_settings_save_id",None): view.frame.after_cancel(view._settings_save_id); view._settings_save_id=None

@contextmanager
def _view_only_traps(app,view):
    hits=[]; originals=[]
    def forbid(label):
        def fail(*_args,**_kwargs): hits.append(label); raise AssertionError(f"view-only callback accessed {label}")
        return fail
    targets=((app,"ws_a_val"),(app,"ws_b_val"),(app,"ws_base_val"),(app,"ws_a_edit"),(app,"ws_b_edit"),(app,"ws_base_edit"),(app,"_request_edit_preload"),(app,"_ensure_edit_loaded"),(app,"_load_edit_workbooks_owned"),(app,"_start_background_thread"),(app,"_atomic_save"),(app,"_atomic_save_with_retry"),(app,"_atomic_replace_file_with_retry"),(app,"_try_alt_save"),(app,"build_manual_b_output_file"),(app,"build_manual_merge_output_file"),(app,"save_a_inplace"),(app,"save_b_inplace"),(app,"save_merged_and_exit"),(view,"refresh"),(view,"_refresh_mode_switch_preserving_selection"),(view,"_start_async_large_only_diff_build"),(view,"_run_copy_action_by_mode"),(view,"_apply_global_sheet_overwrite"),(view,"_apply_selected_column_block"),(sm,"_atomic_save_wb"),(sm,"_align_selected_sheet_snapshots"),(sm,"_compare_selected_sheet_snapshots"))
    try:
        for owner,name in targets:
            if hasattr(owner,name): originals.append((owner,name,getattr(owner,name))); setattr(owner,name,forbid(f"{type(owner).__name__}.{name}"))
        yield hits
    finally:
        for owner,name,original in reversed(originals): setattr(owner,name,original)

@contextmanager
def _publisher_spy(app,view):
    calls=[]; original=view._publish_prepared_cache_surface
    def state():
        entry=dict(app._sheet_exact_entry(view.sheet) or {})
        return {"sheet":view.sheet,"selected":app.selected_sheet,"compute_generation":app._sheet_compute_generation[view.sheet],"exact_generation":entry.get("generation"),"exact_state":entry.get("state"),"ticket":view._mode_switch_seq,"requested":view._mode_switch_requested_value,"pending":view._mode_switch_pending,"building":view._only_diff_async_building}
    def wrapped(*args,**kwargs):
        rows=tuple(kwargs.get("prepared_rows") or ()); record={"before":state(),"rows":rows,"rows_digest":_digest(rows)}; calls.append(record); record["result"]=original(*args,**kwargs); record["after"]=state(); return record["result"]
    view._publish_prepared_cache_surface=wrapped
    try: yield calls
    finally: view._publish_prepared_cache_surface=original

def _assert_one_publish(calls,app,view,action,rows,requested,ticket,compute_generation,exact_generation)->None:
    assert len(calls)==1,(action,calls); call,before=calls[0],calls[0]["before"]
    assert call["result"] is True and call["rows"]==rows and call["rows_digest"]==_digest(rows),(action,call,rows)
    assert before["sheet"]==view.sheet and before["selected"]==view.sheet,(action,before)
    assert before["compute_generation"]==compute_generation and before["exact_generation"]==exact_generation,(action,before)
    assert before["exact_state"]==app._sheet_exact_entry(view.sheet)["state"],(action,before)
    assert before["ticket"]==ticket and before["requested"]==requested and before["pending"] is True and before["building"] is False,(action,before)

def _set_initial_selection(view,pair_idx:int,row_a:int,row_b:int)->None:
    line=view.row_to_line.get(pair_idx); assert line is not None,(pair_idx,view.display_rows)
    view._set_main_selected_cell(line,1); view.selected_pair_idx=pair_idx; view.selected_excel_row=view.selected_excel_row_a=row_a; view.selected_excel_row_b=row_b; view._cursor_cmp_sel_col=view._main_sel_col=1; view._cursor_cmp_sel_line=1; view._update_cursor_lines(); _assert_selection_preserved(view,pair_idx,row_a,row_b)

def _assert_selection_preserved(view,pair_idx:int,row_a:int,row_b:int)->None:
    line=view.row_to_line.get(pair_idx); assert line is not None,(pair_idx,view.display_rows)
    assert view.has_explicit_cell_selection(),(view.selected_pair_idx,view._main_sel_line)
    assert view.selected_pair_idx==pair_idx
    assert view.selected_excel_row==row_a and view.selected_excel_row_a==row_a and view.selected_excel_row_b==row_b
    assert view._main_sel_col==1 and view._cursor_cmp_sel_col==1 and view._main_sel_line==line and view._cursor_cmp_sel_line==1

def _assert_initial_full_virtual(view,full_rows:tuple[int,...],pair_idx:int)->None:
    assert tuple(view._full_display_rows)==full_rows
    assert len(full_rows)>sm._VIRTUAL_VIEWPORT_MAX_ROWS and view._virtual_mode_active()
    cap=min(sm._VIRTUAL_VIEWPORT_MAX_ROWS,len(full_rows)); assert view._virtual_window_start==0; assert tuple(view.display_rows)==full_rows[:cap]; assert pair_idx in view.display_rows,(pair_idx,cap,view.display_rows[:3])

def _assert_disabled_full_virtual(view,full_rows:tuple[int,...],pair_idx:int)->None:
    assert tuple(view._full_display_rows)==full_rows
    assert len(full_rows)>sm._VIRTUAL_VIEWPORT_MAX_ROWS and view._virtual_mode_active()
    cap=min(sm._VIRTUAL_VIEWPORT_MAX_ROWS,len(full_rows)); selected_index=full_rows.index(pair_idx); expected_start=max(0,min(selected_index,max(0,len(full_rows)-cap)))
    assert view._virtual_window_start==expected_start; assert tuple(view.display_rows)==full_rows[expected_start:expected_start+cap]; assert pair_idx in view.display_rows

def _assert_base_insert_contract(view,pair_idx:int,row_a:int,row_b:int)->None:
    """Exact Base-none contract, independent from fallback helper behavior."""
    assert pair_idx not in view.pair_base_row_override
    assert tuple(str(x) for x in view.pair_raw_parts_a[pair_idx])==(str(_INSERT_ID),f"value-{_INSERT_ID}")
    assert tuple(view.pair_raw_parts_b[pair_idx])==tuple(view.pair_raw_parts_a[pair_idx])
    assert not view.pair_diff_cols.get(pair_idx) and view.pair_base_diff_cols.get(pair_idx)=={-1} and view._pair_has_visual_diff(pair_idx)
    assert pair_idx in view.pair_raw_parts_base
    # Side maps may legally retain a physical fallback; that is diagnostic, not
    # the oracle.  Derive the expected prepared Base bytes without helper use.
    fallback=view.mine_to_base_row.get(row_a)
    if fallback is None: fallback=view.theirs_to_base_row.get(row_b)
    expected=("","") if fallback is not None else ("【此侧缺行】","")
    raw=tuple(view.pair_raw_parts_base[pair_idx])
    assert raw==expected,(raw,expected,fallback,dict(view.mine_to_base_row),dict(view.theirs_to_base_row))
    assert tuple(view._prepared_value_for_logical_cell(pair_idx,"BASE",logical_col) for logical_col in (1,2))==raw
    neighbor_row=_INSERT_ID-1+_SCHEMA_ROW_COUNT; neighbor_pair=view.row_a_to_pair_idx.get(neighbor_row)
    assert neighbor_pair is not None and view.row_b_to_pair_idx[neighbor_row]==neighbor_pair
    assert view.pair_base_row_override.get(neighbor_pair)==neighbor_row
    neighbor_raw=tuple(view.pair_raw_parts_base[neighbor_pair])
    assert neighbor_raw==(str(_INSERT_ID-1),f"value-{_INSERT_ID-1}")
    assert not view.pair_base_diff_cols.get(neighbor_pair)
    assert tuple(view._prepared_value_for_logical_cell(neighbor_pair,"BASE",logical_col) for logical_col in (1,2))==neighbor_raw

def _assert_input_hashes(paths,expected)->None: assert paths and {name:_sha256(path) for name,path in paths.items()}==expected
def _shutdown(app)->None:
    if app is not None: app._shutdown_root()
def _assert_temp_settings_path(actual,expected)->None: assert str(actual)==str(expected),(actual,expected)
def _assert_user_settings(path,expected)->None: assert _path_snapshot(path)==expected,path

def _run_case()->None:
    original_settings_path=sm._SETTINGS_PATH; user_settings=Path(original_settings_path); user_before=_path_snapshot(user_settings)
    app=formula_scheduler=root_path=primary_error=None; input_paths={}; input_before={}
    try:
        with tempfile.TemporaryDirectory(prefix="sow_3way_base_insert_") as raw_root:
            root_path=Path(raw_root); base,mine,theirs,merged=(root_path/name for name in ("base.xlsx","mine.xlsx","theirs.xlsx","merged.xlsx"))
            _make_book(base,False); _make_book(mine,True); _make_book(theirs,True); input_paths={"base":base,"mine":mine,"theirs":theirs}; input_before={name:_sha256(path) for name,path in input_paths.items()}
            temp_settings=root_path/"settings.json"; temp_settings.write_text(json.dumps({"only_diff":0}),encoding="utf-8"); sm._SETTINGS_PATH=str(temp_settings); deadline=time.monotonic()+90.0
            print("SMOKE_3WAY_BASE_INSERT_STAGE open-current-exact",flush=True)
            try:
                app,formula_scheduler=_new_app(mine,theirs,base,merged); view=app.sheet_views[_SHEET]
                _wait(app,lambda:(app.selected_sheet==_SHEET and app._is_sheet_exact_current(_SHEET) and app._sheet_exact_entry(_SHEET).get("full_detail_terminal") and view._prepared_complete and view._data_ready and view._row_model_exact and not view._pending_exact_render and view.only_diff_cb.cget("state")=="normal" and not view.only_diff_var.get()),deadline,"selected exact full immutable surface")
                assert time.monotonic()<=deadline and view._is_exact_immutable_view_ready(); assert view._derive_lifecycle_state()=="EDIT_DEFERRED" and not app._edit_workbooks_ready() and not merged.exists()
                print("SMOKE_3WAY_BASE_INSERT_STAGE immutable-base-none",flush=True)
                full_rows=tuple(range(len(view.row_pairs))); assert tuple(view._full_display_rows)==full_rows and len(full_rows)==_ROW_COUNT+_SCHEMA_ROW_COUNT
                row_a=row_b=_INSERT_ID+_SCHEMA_ROW_COUNT; pair_idx=view.row_a_to_pair_idx.get(row_a); assert pair_idx is not None and view.row_b_to_pair_idx[row_b]==pair_idx
                _assert_base_insert_contract(view,pair_idx,row_a,row_b)
                assert view._has_valid_only_diff_snapshot_cache() and tuple(view._only_diff_rows_cache or ())==(pair_idx,)
                _assert_initial_full_virtual(view,full_rows,pair_idx); _set_initial_selection(view,pair_idx,row_a,row_b)
                print("SMOKE_3WAY_BASE_INSERT_STAGE public-enable-cache-only",flush=True)
                before_enable=_hard_snapshot(app,view,input_paths); enable_rows=tuple(view._only_diff_rows_with_touched(view._only_diff_rows_cache)); assert enable_rows==(pair_idx,)
                enable_ticket,compute_generation,exact_generation=view._mode_switch_seq+1,app._sheet_compute_generation[_SHEET],app._sheet_exact_entry(_SHEET)["generation"]
                with _view_only_traps(app,view) as hits,_publisher_spy(app,view) as calls:
                    view.only_diff_cb.invoke(); _wait(app,lambda:(view.only_diff_var.get() and not view._mode_switch_pending and not view._only_diff_async_building and app._is_sheet_exact_current(_SHEET) and app._sheet_exact_entry(_SHEET).get("full_detail_terminal") and view._data_ready and view._prepared_complete and view._only_diff_rows_exact and view._has_valid_only_diff_snapshot_cache() and view.only_diff_cb.cget("state")=="normal"),deadline,"public cache-only enable"); _pump(app)
                assert not hits,hits; _assert_same(before_enable,_hard_snapshot(app,view,input_paths),"public enable plus Tk turn"); _assert_one_publish(calls,app,view,"enable",enable_rows,1,enable_ticket,compute_generation,exact_generation); _assert_selection_preserved(view,pair_idx,row_a,row_b); assert tuple(view._full_display_rows)==enable_rows and tuple(view.display_rows)==enable_rows
                print("SMOKE_3WAY_BASE_INSERT_STAGE public-disable-cache-only",flush=True)
                before_disable=_hard_snapshot(app,view,input_paths); disable_ticket,compute_generation,exact_generation=view._mode_switch_seq+1,app._sheet_compute_generation[_SHEET],app._sheet_exact_entry(_SHEET)["generation"]
                with _view_only_traps(app,view) as hits,_publisher_spy(app,view) as calls:
                    view.only_diff_cb.invoke(); _wait(app,lambda:(not view.only_diff_var.get() and not view._mode_switch_pending and not view._only_diff_async_building and app._is_sheet_exact_current(_SHEET) and app._sheet_exact_entry(_SHEET).get("full_detail_terminal") and view._data_ready and view._prepared_complete and view.only_diff_cb.cget("state")=="normal"),deadline,"public cache-only disable"); _pump(app)
                assert not hits,hits; _assert_same(before_disable,_hard_snapshot(app,view,input_paths),"public disable plus Tk turn"); _assert_one_publish(calls,app,view,"disable",full_rows,0,disable_ticket,compute_generation,exact_generation); _assert_selection_preserved(view,pair_idx,row_a,row_b); _assert_disabled_full_virtual(view,full_rows,pair_idx)
                assert time.monotonic()<=deadline and not app._edit_workbooks_ready() and not merged.exists(); _assert_input_hashes(input_paths,input_before)
                print("SMOKE_3WAY_BASE_INSERT_CASE_OK "+json.dumps({"case":_CASE,"deadline_seconds":90,"input_sha256":input_before,"insert_pair":pair_idx,"only_diff_rows":enable_rows,"full_rows":len(full_rows),"enable_ticket":enable_ticket,"disable_ticket":disable_ticket},sort_keys=True),flush=True)
            except BaseException as exc:
                primary_error=exc; raise
            finally:
                cleanup_errors=[]
                def checked(label,callback):
                    try: callback()
                    except BaseException as exc: cleanup_errors.append((label,exc))
                checked("input SHA before shutdown",lambda:_assert_input_hashes(input_paths,input_before)); checked("temporary settings path",lambda:_assert_temp_settings_path(sm._SETTINGS_PATH,temp_settings)); checked("cancel debounce",lambda:_cancel_debounce(app)); checked("shutdown",lambda:_shutdown(app)); checked("input SHA after shutdown",lambda:_assert_input_hashes(input_paths,input_before)); checked("user settings unchanged",lambda:_assert_user_settings(user_settings,user_before))
                if formula_scheduler is not None: sm.SowMergeApp._schedule_formula_cache_prompt=formula_scheduler
                if cleanup_errors:
                    text="; ".join(f"{label}: {type(error).__name__}: {error}" for label,error in cleanup_errors)
                    if primary_error is not None: primary_error.add_note(f"cleanup secondary failures: {text}")
                    else: raise AssertionError(f"cleanup failures: {text}")
    finally:
        sm._SETTINGS_PATH=original_settings_path
        try:
            _assert_user_settings(user_settings,user_before)
            if root_path is not None: assert not root_path.exists(),root_path
        except BaseException as exc:
            if primary_error is not None: primary_error.add_note(f"outer cleanup secondary failure: {type(exc).__name__}: {exc}")
            else: raise

def main(argv=None)->None:
    parser=argparse.ArgumentParser(); parser.add_argument("--list-cases",action="store_true"); parser.add_argument("--case",choices=(_CASE,)); args=parser.parse_args(argv)
    if args.list_cases:
        if args.case: parser.error("--list-cases cannot be combined with --case")
        print(_CASE,flush=True); return
    selected=(args.case,) if args.case else (_CASE,)
    for case in selected: assert case==_CASE; print(f"SMOKE_3WAY_BASE_INSERT_CASE_START {case}",flush=True); _run_case()
    print(f"SMOKE_3WAY_BASE_INSERT_SUITE_OK ({len(selected)} cases)",flush=True)

if __name__=="__main__": main()
