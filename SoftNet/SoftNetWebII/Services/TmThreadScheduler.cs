using Base.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;

namespace SoftNetWebII.Services
{
    // Thread管理器, 會優先以現有的 Thread 來工作
    // 但主要應用是取代短時間用途的Task使用, 或有需要丟背景或多執行緒 // 若是長時間運行的Thread, 不適合使用
    class SNThreadScheduler : IDisposable
    {
        //// Instance
        private static readonly SNThreadScheduler _instance = new SNThreadScheduler(10);
        public static SNThreadScheduler Instance { get { return _instance; } }

        private object _locker_ = new object();
        private Dictionary<int, TmThread> _thd_scheduler; // Key:ThreadID, Value:TmThread
        private int _thd_action_id;    // ActionID 0..N
        private int _min_thread_count; // keep minimum thread count
        private bool _is_active;
        private Thread _chk_thread;

        public void Initialize(int count = 10)
        {
            _min_thread_count = count;
        }
        private SNThreadScheduler(int min_thread_count)
        {
            lock (_locker_)
            {
                if (_thd_scheduler == null) _thd_scheduler = new Dictionary<int, TmThread>();
                _thd_scheduler.Clear();
                _thd_action_id = -1;
            }
            _min_thread_count = min_thread_count; // 

            _is_active = true;
            _chk_thread = new Thread(SchedulerThread);
            _chk_thread.Name = "SNThreadScheduler" + _chk_thread.ManagedThreadId;
            _chk_thread.IsBackground = true;
            _chk_thread.Start();
        }
        ~SNThreadScheduler()
        {
            Dispose(true);
            //SLogger.log("~SNThreadScheduler");
        }
        // clean up resource
        private bool disposed = false;
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                // If disposing equals true, dispose all managed and unmanaged resources.
                if (disposing)
                {
                    // free managed objects
                    _is_active = false;
                    SpinWait.SpinUntil(() => _chk_thread == null, 500); // wait thread exit

                    // dispose all thread?
                    if (_thd_scheduler != null)
                    {
                        foreach (TmThread t in _thd_scheduler.Values)
                        {
                            if (t != null) t.Dispose(); // dispose this TmThread
                        }
                        _thd_scheduler.Clear();
                    }
                    _thd_scheduler = null;
                    _thd_action_id = -1;

                    if (IsThreadStopped(_chk_thread) == false)
                    {
                        try { _chk_thread.Abort(); }
                        catch { _chk_thread = null; }
                    }
                }
                // free unmanaged objects

                //
                disposed = true;
            }
        }

        // __Thread__
        private void SchedulerThread() // 清Thread (20s)
        {
            while (_is_active == true && !_Fun.Is_Thread_ForceClose)
            {
                try
                {
                    SpinWait.SpinUntil(() => disposed == true || _is_active == false, 20000);
                    if (disposed == true || _is_active == false) break;

                    List<TmThread> need_to_dispose = new List<TmThread>();
                    lock (_locker_)
                    {
                        List<int> AbortTid = new List<int>(); // Abort
                        List<int> IdleAbortTid = new List<int>(); // IdleAbort

                        // check all thread state
                        if (_thd_scheduler == null) _thd_scheduler = new Dictionary<int, TmThread>();
                        foreach (int k in _thd_scheduler.Keys)
                        {
                            TmThread t = _thd_scheduler[k];
                            if (t == null) continue;
                            switch (t.State())
                            {
                                case TmThreadState.Abort: // Abort 要刪
                                    AbortTid.Add(k);
                                    break;
                                case TmThreadState.IdleToAbort: // IdleAbort 不一定要刪
                                    IdleAbortTid.Add(k);
                                    break;
                            }
                        }
                        for (int i = 0; i < AbortTid.Count; i++) // Abort
                        {
                            if (_thd_scheduler != null)
                            {
                                need_to_dispose.Add(_thd_scheduler[AbortTid[i]]);
                                _thd_scheduler.Remove(AbortTid[i]);
                            }
                        }
                        for (int i = 0; i < IdleAbortTid.Count; i++) // IdleAbort
                        {
                            if (_thd_scheduler != null && _thd_scheduler.Count > _min_thread_count) // 保留最少
                            {
                                need_to_dispose.Add(_thd_scheduler[IdleAbortTid[i]]);
                                _thd_scheduler.Remove(IdleAbortTid[i]);
                            }
                            else break;
                        }
                    }
                    // dispose TmThread (不要在 lock 內做 dispose, 會影響使用)
                    for (int i = 0; i < need_to_dispose.Count; i++)
                    {
                        TmThread t = need_to_dispose[i];
                        if (t != null)
                        {
                            //TLogger.Instance.log("TmThread Dispose " + t.UniID().ToString("X16"), "ThreadScheduler");
                            t.Dispose();
                        }
                    }
                    int tc = Count();
                    if (tc > 20)
                    {
                        //TLogger.Instance.log("ThreadCount: " + tc, "ThreadScheduler");
                    }
                }
                catch (Exception ex)
                {
                    //###???SoftNetService.Program._NLogMain.Write_ConfigError(0, "Thread_Scheduler", LogSourceName.Null, "LogDB", RMSError.Service_ThreadScheduler_Fail, 0, ex.Message);
                }
            }
            _chk_thread = null;
        }
        private bool remove(int tid)
        {
            lock (_locker_)
            {
                if (_thd_scheduler != null && _thd_scheduler.ContainsKey(tid) == true)
                {
                    TmThread t = _thd_scheduler[tid];
                    if (t != null) t.Dispose(); // dispose
                    _thd_scheduler[tid] = null;
                    _thd_scheduler.Remove(tid);
                    return true;
                }
                return false;
            }
        }
        private int actionid()
        {
            lock (_locker_)
            {
                _thd_action_id++;
                if (_thd_action_id < 0 || _thd_action_id >= Int32.MaxValue) _thd_action_id = 0;
                return _thd_action_id;
            }
        }

        /// <summary>
        /// start new delegate Action with timeout // 會優先使用idle thread,不足時才會新建
        /// </summary>
        /// <param name="item"></param>
        /// <param name="timeout">ms, no timeout is -1</param>
        /// <param name="wait">False NoWait, True Wait</param>
        /// <returns>
        /// if wait = false, return Key ThreadID+ActionID, for Wait(tid+aid) or Check(tid+aid) // 高32bit是tid,低32bit是aid
        /// if wait = true,  return Action result (with timeout or abort?)
        /// 0xFF66  Action item is null
        /// 0xFF99  error 建不出來
        /// </returns>
        public long StartAction(Action item, int timeout = -1, bool wait = false) // default 沒Timeout,不等待這行做完
        {
            // Get call stack
            ////StackTrace stackTrace = new StackTrace();
            // Get calling method name
            ////TLogger.Instance.log("StartAction by " + stackTrace.GetFrame(1).GetMethod().Name, "ThreadScheduler");

            if (item == null) return -154; // null item, no need // 0xFF66
            long uid = -1;
            int i = 0, retry = 10;
            do
            {
                uid = -1;
                try
                {
                    //TLogger.Instance.log("start action 1", "ThreadScheduler");
                    lock (_locker_)
                    {
                        // check all thread state
                        if (_thd_scheduler == null) _thd_scheduler = new Dictionary<int, TmThread>();
                        foreach (TmThread t in _thd_scheduler.Values)
                        {
                            if (t == null) continue;
                            TmThreadState st = t.State();
                            ////TLogger.Instance.log(t.UniID().ToString("X16") + "-" + _thd_scheduler.Values.Count + " is " + st.ToString() + " " + t._timeout.ToString(), "ThreadScheduler");
                            if (st == TmThreadState.Idle || st == TmThreadState.IdleToAbort) // 有idle thread 可使用
                            {
                                ////TLogger.Instance.log("start action 1-1", "ThreadScheduler");
                                t.SetAction(item, actionid(), timeout); // 會先跑
                                uid = t.UniID();
                                ////TLogger.Instance.log("start action 1-2", "ThreadScheduler");
                                break; // exit foreach
                            }
                        }
                    }
                    if (uid < 0) // not found idle thread
                    {
                        //TLogger.Instance.log("start action 2", "ThreadScheduler");
                        // 沒有idle thread, 則建一個新的
                        TmThread t = new TmThread();
                        int tid = t.ThdID();
                        lock (_locker_)
                        {
                            if (tid >= 0 && _thd_scheduler != null && _thd_scheduler.ContainsKey(tid) == false)
                            {
                                //TLogger.Instance.log("start action 2-1", "ThreadScheduler");
                                _thd_scheduler.Add(tid, t);
                                t.SetAction(item, actionid(), timeout); // 會先跑
                                uid = t.UniID();
                                //TLogger.Instance.log("start action 2-2", "ThreadScheduler");
                                break; // exit do while
                            }
                        }
                        // 建不出來或無法加入dict? (若可以,在上面就會break)
                        uid = -1;
                        if (t != null) t.Dispose(); // error
                        t = null;
                    }
                }
                catch (Exception ex)
                {
                    //TLogger.Instance.log(i.ToString() + " TmThread Exception: " + ex.Message, "ThreadScheduler");
                    //###???SoftNetService.Program._NLogMain.Write_ConfigError(0, "Thread_Scheduler", LogSourceName.Null, "LogDB", RMSError.Service_ThreadScheduler_Fail, 0, ex.Message);
                }
                if (uid >= 0) break; // exit do while
                // uid < 0, retry
                if (++i < retry) Thread.Sleep(40);
                else
                {
                    //###???SoftNetService.Program._NLogMain.Write_RunError(0, "", "Thread_Scheduler", RMSError.Service_ThreadScheduler_Fail, LogSourceName.Null, "", 0, "Thread 建不出來.");
                    return -103; // Error,建不出來! // 0xFF99
                }
                if (_Fun.Is_Thread_ForceClose) {  break; }
            } while (true);
            ////TLogger.Instance.log("start action 3 " + uid.ToString("X"), "ThreadScheduler");
            // wait or return
            /*
            if (wait == false)
            {
                if (uid < 0)
                {
                    SoftNetService.Program._LOG.Write_RunError(0, "", "Thread_Scheduler", RMSError.Service_ThreadScheduler_Fail, LogSourceName.Null, "", 0, ToolFun.StringAdd("error code=", uid.ToString()));
                }
                return uid;
            }
            */
            return Wait(uid, timeout);
        }
        // 0  Running  還在做
        // 1  Done     做完
        // -1 Timeout  逾時
        // -2 Abort    此ThreadID將退出
        // -3 Error    此ThreadID不存在
        // -4 Unknown  ??
        // -5 Exception
        public int Wait(long uid, int timeout = -1) // 卡住
        {
            int re = 0;
            try
            {
                SpinWait.SpinUntil(() => (re = Check(uid)) != 0, timeout); // wait Action
                return re; // 0 SpinUntil timeout
            }
            catch { return re; }
        }
        public int Wait(params long[] uid) // 要全部都Done
        {
            int re = 0;
            try
            {
                if (uid == null || uid.Length <= 0) return 1; // done
                SpinWait.SpinUntil(() =>
                {
                    for (int i = 0; (uid != null && i < uid.Length); i++)
                    {
                        if ((re = Check(uid[i])) == 0) return false; // still doing
                    }
                    return true; // all done
                }, -1); // wait Action
                return re; // 0 SpinUntil timeout
            }
            catch { return re; }
        }
        public int Check(long uid) // 一次性
        {
            try
            {
                if (_is_active == false) return -16; // dispose

                int tid = (int)(uid >> 32);        // High32bit
                int aid = (int)(uid & 0xFFFFFFFF); // Low32bit
                lock (_locker_)
                {
                    TmThread t = null;
                    if (_thd_scheduler == null) _thd_scheduler = new Dictionary<int, TmThread>();
                    if (_thd_scheduler.ContainsKey(tid) == true) t = _thd_scheduler[tid];

                    if (t == null) return -3; // error
                    switch (t.State(aid))
                    {
                        case TmThreadState.Idle: return 1; // done
                        case TmThreadState.Run: return 0; // running
                        case TmThreadState.Timeout: return -1; // timeout
                        case TmThreadState.Abort: return -2; // abort
                        case TmThreadState.IdleToAbort: return -2; // abort
                        default: return -4; // unknown?
                    }
                }
            }
            catch { return -5; }
        }
        public int Count()
        {
            lock (_locker_)
            {
                if (_thd_scheduler == null) _thd_scheduler = new Dictionary<int, TmThread>();
                return _thd_scheduler.Count;
            }
        }

        enum TmThreadState : byte
        {
            Idle = 0,
            Run = 1,
            Timeout = 2, // Run to Timeout
            Abort = 3,   // Run to Abort
            IdleToAbort = 4, // Idle to Abort // 發呆太久的Thread,是否刪除?
            LongToAbort = 5, // Run(Long) to Abort // 真的跑太久的Thread,是否刪除?(兩小時?) // Thread Count真的太多時，才會啟動這個機制?
        };
        // sub-class
        class TmThread : IDisposable // 只做事(InvokeAction),是否Timeout或是要Abort,再由外部的Thread來處理
        {
            private object _tlocker_ = new object();
            private Stopwatch _sw;  // time for idle or running
            private Thread _invoke; // thread to action invoke
            private Action _action; // delegate
            private int _actid;     // action id
            private int _timeout;
#if true // Release
            private const int DEFAULT_RUNABORT = 180000;  // Run  至少180s後才會RunToAbort (依使用者設的_timeout*4倍時間,且至少要180s)
            private const int DEFAULT_IDLEABORT = 300000; // Idle 300s後才會IdleToAbort
            private const int DEFAULT_LONGABORT = 7200000;// 7200s
#else // Test
            private const int DEFAULT_RUNABORT = 20000;
            private const int DEFAULT_IDLEABORT = 40000;
#endif
            public TmThread()
            {
                _actid = -1;
                action = null;

                _invoke = new Thread(InvokeThread);
                _invoke.Name = "TmThread" + _invoke.ManagedThreadId;
                _invoke.IsBackground = true;
                _invoke.Start();
            }
            // __Thread__
            private void InvokeThread() // 做事的Thread, 主要處理_action
            {
                while (disposed == false)
                {
                    // 這裡不要 try-catch 因為這裡只是跑其他Action (真正Exception的是其他Action)
                    SpinWait.SpinUntil(() => disposed == true || action != null, 60000); // action
                    if (disposed == true) break;
                    if (action == null) continue;
                    // doing
                    action.Invoke(); // 做事,做完才出來
                    // idle
                    action = null; // 做完
                }
                action = null;
                _invoke = null;
            }
            private Action action
            {
                get { return _action; } // 不 locker
                set
                {
                    lock (_tlocker_)
                    {
                        reset();         // 時間重計(因為 sw 是 idle/running 共用,必須先清,免得誤認狀態)
                        _action = value;
                    }
                }
            }
            private void reset()
            {
                lock (_tlocker_)
                {
                    if (_sw == null) _sw = new Stopwatch();
                    _sw.Reset();
                    _sw.Start();
                }
            }

            // ThreadID + ActionID = Unique Identifier
            protected internal bool SetAction(Action item, int aid, int timeout)
            {
                lock (_tlocker_)
                {
                    if (disposed == true || _invoke == null) return false;
                    if (item == null) return false;
                    if (action != null) return false; // action is doing, could not be set
                    if (aid < 0) aid = -aid; //不該小於0

                    _timeout = timeout; // <=0表示沒有timeout,要一直等到做完
                    _actid = aid;  // 要比_action早設
                    action = item; // 最後設,因為設定_action就會動,且狀態改變
                    return true;
                }
            }
            protected internal long UniID()
            {
                long a = ThdID();
                a = a << 32;     // High32bit is ThreadID
                a = a + ActID(); // Low32bit  is ActionID
                return a;
            }
            protected internal int ThdID() // ThreadID
            {
                lock (_tlocker_)
                {
                    if (disposed == true || _invoke == null) return -1;
                    return _invoke.ManagedThreadId;
                }
            }
            protected internal int ActID() // ActionID
            {
                lock (_tlocker_)
                {
                    if (disposed == true || _invoke == null) return -1;
                    if (action == null) return -1;
                    return _actid;
                }
            }
            protected internal TmThreadState State(int aid = -1) // action id // < 0 表示沒指定 action id , 會 check 最後一次 action 狀況
            {
                lock (_tlocker_)
                {
                    if (disposed == true || _invoke == null) return TmThreadState.IdleToAbort;

                    long now = _sw.ElapsedMilliseconds;
                    if (action == null) // Idle, IdleToAbort
                    {
                        if (now > DEFAULT_IDLEABORT) return TmThreadState.IdleToAbort;
                        return TmThreadState.Idle;
                    }
                    else // Run, Timeout, Abort
                    {
                        if (aid >= 0 && aid != _actid) return TmThreadState.Idle; // action id 不符合, 表示此 id 可能已做完, 算Idle

                        if (_timeout > 0)
                        {
                            long AbortTime = (((long)_timeout) << 2); // <<2 = *4 // 1 for run , +3 for abort thread
                            if (now > AbortTime && now > DEFAULT_RUNABORT) return TmThreadState.Abort;
                            else if (now > _timeout) return TmThreadState.Timeout;
                        }
                        return TmThreadState.Run;
                    }
                }
            }

            // clean up resource
            private bool disposed = false;
            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
            protected virtual void Dispose(bool disposing)
            {
                if (!this.disposed)
                {
                    // If disposing equals true, dispose all managed and unmanaged resources.
                    if (disposing)
                    {
                        //if (_invoke != null) Debug.WriteLine(_invoke.ManagedThreadId + "    Exit");
                        disposed = true;
                        // free managed objects
                        SpinWait.SpinUntil(() => _invoke == null, 500); // wait thread exit

                        if (IsThreadStopped(_invoke) == false)
                        {
                            try { _invoke.Abort(); }
                            catch { _invoke = null; }
                        }
                        _sw = null;
                        _actid = -1;
                        _action = null;
                    }
                    // free unmanaged objects

                    //
                    disposed = true;
                }
            }
        } // end of TmThread sub-class

        private static bool IsThreadStopped(Thread t)
        {
            if (t == null) return true; // null is stop

            if (t.IsAlive == false)
            {
                // State = Stopped or Aborted
                if ((t.ThreadState & (System.Threading.ThreadState.Stopped | System.Threading.ThreadState.Aborted)) != 0)
                {
                    return true;
                }
            }
            return false; // thread is no stop
        }
    } // end of SNThreadScheduler

}
