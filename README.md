# API Timing Functions VBA
Microsoft Win32 API Timing Functions in VBA7


| VBA Function                         | Description                                                                                                   |
| ------------------------------------ | --------------------------------------------------------------------------------------------------------------|
| `timestamp()`                        | Returns VBA `Time$` as string suffixed with Milliseconds, e.g. `18:45:00.567`                                 |
| `kernel_sleep(10000)`                | Sleeps for 10000 Milliseconds while keeping VBA responsive. Option to add your code in wait loop.             |
| `get_qpc_seconds()`                  | Returns number of Seconds since last system boot                                                              | 
| `get_qpc_milliseconds()`             | Returns number of Milliseconds since last system boot                                                         |
| `get_qpc_microseconds()`             | Returns number of Microseconds since last system boot                                                         |


<details><summary>Win32 Function References</summary>
<p>

[Query Performance Frequency](https://learn.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancefrequency)  
[Query Performance Counter](https://learn.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancecounter)
  
</p>
</details>  

QPC functions can be used for time-interval measurements, e.g. before and after your code.
