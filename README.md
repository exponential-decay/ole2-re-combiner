# ole2-re-combiner
Jython OLE2 re-Combing tool...

Can be used with Jython with Apache POI on the CLASSPATH. Example command to
run:

* `jython -Dpython.path=$CLASSPATH ole2recombiner.py --dir Dir2Combine/`
* `jython -Dpython.path=$CLASSPATH ole2recombiner.py --ole2 file2split.doc`

Example on Windows:

* `jython -Dpython.path=%CLASSPATH%\poi-3.13-beta1-20150723.jar ole2recombiner.py`

Example CLASSPATH: `:/usr/bin/poi/poi-3.11-beta2/poi-3.11-beta2-20140822.jar`

If you haven't Apache POI on the `CLASSPATH` you can have it in the same folder
as `ole2recombiner.py` and run the following as an example:

* `jython -Dpython.path=poi-4.0.0.jar ole2recombiner.py --combine dir/`

# License

Copyright 2015 Ross Spencer

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
