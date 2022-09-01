//
//  ContentView.swift
//  MSOffice RPC
//
//  Created by Oryon Ivanov on 2/09/22.
//

import SwiftUI
import PythonKit

func runPythonCode() {
    let sys = Python.import("sys")
    sys.path.append("/Users/oryonivanov/Documents/Private Projects/MSOffice DiscordRPC/MSOffice RPC/MSOffice RPC/Python Scripts")
    let word_rpc = Python.import("Word_RPC")
}

struct ContentView: View {
    @State private var show = false
    var body: some View {
        VStack {
            if !show {
                RootView(show: $show)
                    .frame(maxWidth: .infinity, maxHeight: .infinity)
                    .transition(AnyTransition.move(edge: .leading)).animation(.default)
            }
            if show {
                NextView(show: $show)
                    .frame(maxWidth: .infinity, maxHeight: .infinity)
                    .transition(AnyTransition.move(edge: .trailing)).animation(.default)
            }
            
        }
    }
    
    struct RootView: View {
        @Binding var show: Bool
        var body: some View {
            VStack{
                Text(Image(systemName: "hammer.fill"))
                    .foregroundColor(.blue)
                    .padding(10)
                    .font(.largeTitle)
                
                Text("MSOffice RPC")
                    .bold()
                    .font(.largeTitle)
                
                Button(action: {
                    self.show = true
                }, label: {
                    Text("Continue")
                        .frame(width: 200,
                               height: 35,
                               alignment: .center)
                        .background(Color.blue)
                        .foregroundColor(.white)
                        .cornerRadius(8)
                })
                .buttonStyle(.plain)
            }
        }
    }
    
    
    struct NextView: View {
        @Binding var show: Bool
        var body: some View {
            VStack{
                Text(Image(systemName: "hammer.fill"))
                    .foregroundColor(.blue)
                    .padding(10)
                    .font(.largeTitle)
                
                Button(action: {
                    self.show = false
                }, label: {
                    Text("Back")
                        .frame(width: 200,
                               height: 35,
                               alignment: .center)
                        .background(Color.blue)
                        .foregroundColor(.white)
                        .cornerRadius(8)
                })
                .buttonStyle(.plain)
                
                Button(action: {
                    runPythonCode()
                }, label: {
                    Text("testing")
                        .frame(width: 200,
                               height: 35,
                               alignment: .center)
                        .background(Color.blue)
                        .foregroundColor(.white)
                        .cornerRadius(8)
                })
                .buttonStyle(.plain)
            }
        }
    }
    
    struct ContentView_Previews: PreviewProvider {
        static var previews: some View {
            ContentView()
        }
    }
}


