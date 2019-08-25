1. 동작 설명
 - Program.cs의 appInstance는 Excel instance를 담기위한 Dictionary => Open(Create)부터 Close까지 instance가 존재
 - Form1.cs의 Example_Click이 Activity들을 실행시킴
 - 각 Activity들은 원하는 Excel Instance에 접근하기 위해 Program 객체를 받아서 appInstance Dictionary에서 꺼내옴
 - Excel Instance를 appInstance로부터 얻으면 각 Activity들은 자기 역할을 수행

2. 한계
 현재 Example_Click에 있는 것처럼 코드 상으로는 정상적으로 동작이 가능함.
 하지만 하나의 Workflow flowchart에 각 Activity들을 넣어서 실행시키려고 하면 System.NotSupportedException 때문에 실행이 안됨.
 이 부분에 대해서는 해결하지 못한 상태.