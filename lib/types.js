module.exports = {
  /**
   * sheet(table) 的类型
   * 影响输出json的类型
   * 当有#id类型的时候  表输出json的是map形式(id:{xx:1})
   * 当没有#id类型的时候  表输出json的是数组类型 没有id索引
   */
  SheetType: {
    /**
     * 普通表 
     * 输出JSON ARRAY
     */
    NORMAL: 0,

    /**
     * 有主外键关系的主表
     * 输出JSON MAP
     */
    MASTER: 1,

    /**
     * 有主外键关系的从表
     * 输出JSON MAP
     */
    SLAVE: 2
  },

  /**
   * 支持的数据类型
   */
  DataType: {
    INT: 'int',
    LONG: 'long',
    FLOAT: 'float',
    STRING: 'string',
    BOOL: 'bool',
    DATE: 'date',
    ARRAY: '[]',
    ARRAY_INT: 'int[]',
    ARRAY_FLOAT: 'float[]',
    ARRAY_STRING: 'string[]',
    OBJECT: '{}',
    UNKNOWN: 'unknown',
    MAIN_KEY: 'main key'
  },

  /**
   * 支持的转出数据类型 hash格式 / array格式
   */
  JsonType: {
    HASH: "hash",
    ARRAY: "array"
  }
};