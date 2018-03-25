Rails.application.routes.draw do
  root 'rankings#new'
  post '/', to: 'rankings#new'
end
