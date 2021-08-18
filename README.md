# ExcelToLua
Excel转Lua表：提取公共table。
第一种方式通过setmetable来访问。
添加一种setconfigtable来访问，经测试比setmetatable快27倍。需要修改lua源码。

给lua添加一个新的flag：
```
#define TM_INDIRECT_INDEX (TM_EQ + 1)
```
然后修改luaV_finishget：

```
void luaV_finishget (lua_State *L, const TValue *t, TValue *key, StkId val,
                      const TValue *slot) {
  int loop;  /* counter to avoid infinite loops */
  const TValue *tm, *mk;  /* metamethod */
  Table *ot, *mt;
  for (loop = 0; loop < MAXTAGLOOP; loop++) {
    if (slot == NULL) {  /* 't' is not a table? */
      lua_assert(!ttistable(t));
      tm = luaT_gettmbyobj(L, t, TM_INDEX);
      if (ttisnil(tm))
        luaG_typeerror(L, t, "index");  /* no metamethod */
      /* else will try the metamethod */
    }
    else {  /* 't' is a table */
      lua_assert(ttisnil(slot));
      ot = hvalue(t);
      mt = ot->metatable;
      tm = fasttm(L, mt, TM_INDEX);  /* table's metamethod */
      if (tm == NULL) {  /* no metamethod? */

        /*HuaHua*/
        if (mt != NULL && (mt->flags & (1u << TM_INDIRECT_INDEX))){
            mk = luaH_get(mt, key);
            if (mk != NULL){
                slot = luaH_get(ot, mk);
                setobj2s(L, val, slot);
                return;
            }
        }

        setnilvalue(val);  /* result is nil */
        return;
      }
      /* else will try the metamethod */
    }
    if (ttisfunction(tm)) {  /* is metamethod a function? */
      luaT_callTM(L, tm, t, key, val, 1);  /* call it */
      return;
    }
    t = tm;  /* else try to access 'tm[key]' */
    if (luaV_fastget(L,t,key,slot,luaH_get)) {  /* fast track? */
      setobj2s(L, val, slot);  /* done */
      return;
    }
    /* else repeat (tail call 'luaV_finishget') */
  }
  luaG_runerror(L, "'__index' chain too long; possible loop");
}
```
