# API Timing Functions VBA
Microsoft Win32 API Timing Functions in VBA7


| VBA Function                         | Description                                                                                                   |
| ------------------------------------ | --------------------------------------------------------------------------------------------------------------|
| `kernel_sleep(10000)`                | Sleeps for 10000 Milliseconds while keeping VBA responsive. Option to add your code in wait loop.             |
| `timestamp()`                        | Returns VBA Time() as string suffixed with Milliseconds, e.g. `18:45:00.567`                                  |
| `get_qpc_seconds()`                  | Returns number of Seconds since last system boot                                                              | 
| `get_qpc_milliseconds()`             | Returns number of Milliseconds since last system boot                                                         |
| `get_qpc_microseconds()`             | Returns number of Microseconds since last system boot                                                         |
