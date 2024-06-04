# FACIAL-RECOGNITION-ATTENDANCE-SYSTEM
Facial Recognition Attendance System automates attendance tracking using facial recognition technology and also gives voice feedback for the user. It replaces manual methods with an efficient process.

PROBLEM :
Attendance management systems often involves outdated methods that are time-consuming and prone to errors. Traditional techniques like manual tracking create inefficiencies and add to administrative workload. Moreover, biometric systems relying on fingerprints can be impractical in environments where hygiene is critical. To meet technical standards and improve efficiency, there’s a need for a modern, contactless solution for attendance management. Resolving these issues will ensure adherence to technical standards while simplifying and enhancing the speed of attendance tracking processes.

PROPOSED SYSTEM :
The proposed software has mainly two processes. First, when the subject name and subject duration is given, the camera will turn ON and start recording video frames. Then using Open CV and face recognition library of Python , the system encodes the current video frames and gives face encodings and saves it. There will be reference datas which have various images of people who need to be included in this attendance system which are added with our needs according to the number of persons and these faces will then be encoded as known face encodings. Secondly, the known face encodings and face encodings (received from video frames) will be compared and if a successful matching is found then the person is successfully detected. When a successful detection is happens the attendance of the person is marked in the sheet with time at which the person is detected. When name is detected and recorded into the excel sheet, a voice feedback is given, to know if their attendance is correctly marked into the sheet. For giving this feedback pyttsx3 library of Python is used.


NOTE :
Install Python (3.9 recommended) and Dlib, openCV, face_recognition , openpyxl libraries.
